using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.ExtensionManager.UI;

namespace ExtensionSync
{
    public class ExtensionManagerFacade
    {

        public ExtensionManagerFacade(IVsExtensionManager visualStudioExtensionManager,
                                      IVsExtensionRepository visualStudioExtensionRepository)
        {
            ExtensionManager = visualStudioExtensionManager;
            ExtensionRepository = visualStudioExtensionRepository;
        }

        public event Action<string> Log;

        public List<ExtensionInformation> GetInstalledExtensionsInformation()
        {
            var installedExtensions = ExtensionManager.GetInstalledExtensions();
            var userExtensions = installedExtensions.Where(ext =>
                        !ext.Header.SystemComponent && !ext.Header.InstalledByMsi)
                        .OrderBy(ext => ext.Header.Name);

            return userExtensions.Select(
                    e => new ExtensionInformation
                    {
                        Name = e.Header.Name,
                        Identifier = e.Header.Identifier,
                        Version = e.Header.Version
                    })
                    .ToList();
        }

        public void InstallExtensions(List<ExtensionInformation> extensions)
        {
            if (extensions == null || extensions.Count == 0)
                return;

            try
            {
                foreach (var extension in extensions)
                {
                    var query = ExtensionRepository.CreateQuery<VSGalleryEntry>(false, true)
                                .OrderByDescending(v => v.Ranking)
                                .Skip(0)
                                .Take(25) as IVsExtensionRepositoryQuery<VSGalleryEntry>;
                    if (query == null) continue;

                    query.ExecuteCompleted += QueryExecuteCompleted;
                    query.SearchText = extension.Name;
                    query.ExecuteAsync(extension);
                }
            }
            catch (Exception exception)
            {
                LogMessage(string.Format("Error while installing extensions: {0}", exception.Message));
            }
        }

        public void UnInstallExtensions(List<ExtensionInformation> extensions, DateTimeOffset configUpdateDateTime)
        {
            if (extensions == null || extensions.Count == 0)
                return;

            var installedUserExtensions = ExtensionManager.GetInstalledExtensions().
                                    Where(e => !e.Header.SystemComponent).ToList();

            var extensionsInstalledAfterConfigUpdated = installedUserExtensions
                            .Where(i => i.InstalledOn > configUpdateDateTime).ToList();

            foreach (var extension in extensions)
            {
                try
                {
                    var extensionInformation = extension;
                    var userExtension =
                        installedUserExtensions.SingleOrDefault(e =>
                            e.Header.Name == extensionInformation.Name &&
                            e.Header.Identifier == extensionInformation.Identifier);
                    if (userExtension == null) continue;

                    if (extensionsInstalledAfterConfigUpdated.Contains(userExtension))
                    {
                        LogMessage(string.Format("Not uninstalling {0} since it was installed after last update to config file."
                            , userExtension.Header.Name));
                        continue;
                    }

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
                var extensionName = extensionInformation.Name;
                var entry =
                    e.Results.Cast<VSGalleryEntry>().SingleOrDefault(r =>
                        r.Name == extensionInformation.Name && r.VsixID == extensionInformation.Identifier);
                if (entry == null)
                {
                    LogMessage(string.Format("Could not find {0} in Online Repository", extensionName));
                    return;
                }

                var installedExtensions = GetInstalledExtensionsInformation();
                var installedExtension = installedExtensions.FirstOrDefault(ext => ext.Name == extensionName
                                    &&
                                    ext.Identifier ==
                                    extensionInformation.Identifier);
                if (AutoUpdateExtensions)
                {
                    if (installedExtension != null && installedExtension.Version.ToString() == entry.VsixVersion)
                        return;
                }

                try
                {
                    var installableExtension = ExtensionRepository.Download(entry);
                    if (installableExtension == null)
                        return;

                    if (installedExtension != null)
                    {
                        if (installedExtension.Version >= installableExtension.Header.Version)
                            return;
                        //extension needs to be updated - uninstall and install again
                        LogMessage(string.Format("{0} has an update available.", extensionName));
                        UnInstallExtensions(new List<ExtensionInformation> { installedExtension }, DateTime.Now);
                    }

                    ExtensionManager.Install(installableExtension, false);
                    LogMessage(string.Format("Installed {0}", extensionName));
                }
                catch (Exception exception)
                {
                    LogMessage(string.Format("Error while installing {0}: {1}", extensionName, exception.Message));
                }
            }
            catch (Exception exception)
            {
                LogMessage(string.Format("Error while installing extensions: {0}", exception.Message));
            }
        }

        void LogMessage(string message)
        {
            if (Log != null)
                Log.Invoke(message);
        }

        public bool AutoUpdateExtensions { get; set; }
        IVsExtensionManager ExtensionManager { get; set; }
        IVsExtensionRepository ExtensionRepository { get; set; }
    }
}
