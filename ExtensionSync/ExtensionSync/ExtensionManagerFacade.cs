using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.ExtensionManager.UI;

namespace LatishSehgal.ExtensionSync
{
    public class ExtensionManagerFacade
    {

        public ExtensionManagerFacade(IVsExtensionManager visualStudioExtensionManager,
                                      IVsExtensionRepository visualStudioExtensionRepository)
        {
            ExtensionManager = visualStudioExtensionManager;
            ExtensionRepository = visualStudioExtensionRepository;

            processingTimer.Elapsed += (sender, e) => ClearExtensionManagerEventHandlers();
        }

        public event Action<string> Log;

        public List<ExtensionInformation> GetInstalledExtensionsInformation()
        {
            var installedExtensions = ExtensionManager.GetInstalledExtensions();
            var userExtensions = installedExtensions.Where(ext => 
                        !ext.Header.SystemComponent && !ext.Header.InstalledByMsi)
                        .OrderBy(ext => ext.Header.Name);
            return userExtensions.Select(
                    e => new ExtensionInformation { Name = e.Header.Name, Identifier = e.Header.Identifier })
                    .ToList();
        }

        public void InstallExtensions(List<ExtensionInformation> extensions)
        {
            if (extensions == null || extensions.Count == 0)
                return;

            try
            {
                extensionsBeingInstalled.Clear();
                ExtensionRepository.DownloadCompleted += ExtensionRepositoryDownloadCompleted;
                ExtensionManager.InstallCompleted += ExtensionManagerInstallCompleted;

                foreach (var extension in extensions)
                {
                    extensionsBeingInstalled.Add(extension.Name);

                    var query = ExtensionRepository.CreateQuery<VSGalleryEntry>(false, true)
                                .OrderByDescending(v => v.Ranking)
                                .Skip(0)
                                .Take(25) as IVsExtensionRepositoryQuery<VSGalleryEntry>;
                    if (query == null) continue;

                    query.ExecuteCompleted += QueryExecuteCompleted;
                    query.SearchText = extension.Name;
                    query.ExecuteAsync(extension);
                }

                //Give it a a few minutes to do its work. If everything goes smoothly, event handlers should unbind.
                // Make a safety check and remove handlers if not already done
                processingTimer.Start();
            }
            catch (Exception exception)
            {
                LogMessage(string.Format("Error while installing extensions: {0}", exception.Message));
            }
        }

        private void ClearExtensionManagerEventHandlers()
        {
            lock (this)
            {
                if (eventHandlersCleared) return;

                eventHandlersCleared = true;
                ExtensionRepository.DownloadCompleted -= ExtensionRepositoryDownloadCompleted;
                ExtensionManager.InstallCompleted -= ExtensionManagerInstallCompleted;

                processingTimer.Stop();
            }
        }

        public void UnInstallExtensions(List<ExtensionInformation> extensions)
        {
            if (extensions == null || extensions.Count == 0)
                return;

            var installedUserExtensions = ExtensionManager.GetInstalledExtensions().
                                    Where(e => !e.Header.SystemComponent).ToList();
            foreach (var extension in extensions)
            {
                try
                {
                    if (extension.Name.Equals("ExtensionSync", StringComparison.InvariantCultureIgnoreCase))
                        continue;

                    var extensionInformation = extension;
                    var userExtension =
                        installedUserExtensions.SingleOrDefault(e =>
                            e.Header.Name == extensionInformation.Name &&
                            e.Header.Identifier == extensionInformation.Identifier);
                    if (userExtension == null) continue;

                    LogMessage(string.Format("Uninstalling {0}", userExtension.Header.Name));
                    ExtensionManager.Uninstall(userExtension);
                }
                catch (Exception exception)
                {
                    LogMessage(string.Format("Error while uninstalling {0}: {1}", extension.Name, exception.Message));
                }
            }
        }

        private void QueryExecuteCompleted(object sender, ExecuteCompletedEventArgs e)
        {
            try
            {
                if (e.Error != null)
                {
                    LogMessage(string.Format("Error while searching online for extension: {0}", e.Error.Message));
                    return;
                }

                var extensionInformation = (ExtensionInformation)e.UserState;
                var entry =
                    e.Results.Cast<VSGalleryEntry>().SingleOrDefault(r => 
                        r.Name == extensionInformation.Name && r.VsixID == extensionInformation.Identifier);
                if (entry == null)
                {
                    LogMessage(string.Format("Could not find {0} in Online Repository", extensionInformation.Name));
                    ExtensionInstallDone(extensionInformation.Name);
                    return;
                }
                ExtensionRepository.DownloadAsync(entry);
            }
            catch (Exception exception)
            {
                LogMessage(string.Format("Error while installing extensions: {0}", exception.Message));
            }
        }

        private void ExtensionRepositoryDownloadCompleted(object sender, DownloadCompletedEventArgs e)
        {
            try
            {
                if (e.Error != null)
                {
                    LogMessage(string.Format("Error while downloading extension: {0}", e.Error.Message));
                    return;
                }

                var installableExtension = e.Payload;

                var installedExtensions = GetInstalledExtensionsInformation();
                if (installedExtensions.Any(i =>
                        i.Name == installableExtension.Header.Name &&
                        i.Identifier == installableExtension.Header.Identifier))
                {
                    ExtensionInstallDone(installableExtension.Header.Name);
                    return;
                }

                if (installableExtension == null)
                    return;
                ExtensionManager.InstallAsync(installableExtension, false);
            }
            catch (Exception exception)
            {
                LogMessage(string.Format("Error while installing extensions: {0}", exception.Message));
            }
        }

        private void ExtensionManagerInstallCompleted(object sender, InstallCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                LogMessage(string.Format("Error while installing extension: {0}", e.Error.Message));
                return;
            }
            var extensionName = e.Extension.Header.Name;
            LogMessage(string.Format("Installed {0}", extensionName));

            ExtensionInstallDone(extensionName);
            CheckIfAllExtensionInstallsHaveCompleted();
        }

        private void CheckIfAllExtensionInstallsHaveCompleted()
        {
            lock (this)
            {
                if (extensionsBeingInstalled.Count == 0)
                    ClearExtensionManagerEventHandlers();
            }

        }

        private void ExtensionInstallDone(string extensionName)
        {
            lock (this)
            {
                extensionsBeingInstalled.Remove(extensionName);
            }
        }

        void LogMessage(string message)
        {
            if (Log != null)
                Log.Invoke(message);
        }

        IVsExtensionManager ExtensionManager { get; set; }
        IVsExtensionRepository ExtensionRepository { get; set; }

        List<string> extensionsBeingInstalled = new List<string>();
        private bool eventHandlersCleared;
        private Timer processingTimer = new Timer((MaxProcessingDuration));

        private const int MaxProcessingDuration = 120000;
    }
}
