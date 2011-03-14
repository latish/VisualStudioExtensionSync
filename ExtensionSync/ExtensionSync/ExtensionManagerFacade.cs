using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.ExtensionManager.UI;

namespace LatishSehgal.ExtensionSync
{
    public class ExtensionManagerFacade
    {
        public ExtensionManagerFacade(IVsExtensionManager visualStudioExtensionManager, IVsExtensionRepository visualStudioExtensionRepository)
        {
            ExtensionManager = visualStudioExtensionManager;
            ExtensionRepository = visualStudioExtensionRepository;

            ExtensionRepository.DownloadCompleted += ExtensionRepository_DownloadCompleted;
            ExtensionManager.InstallCompleted += ExtensionManager_InstallCompleted;
        }

        public event Action<string> Log;

        public List<ExtensionInformation> GetInstalledExtensionsInformation()
        {
            var installedExtensions = ExtensionManager.GetInstalledExtensions();
            var userExtensions = installedExtensions.Where(ext => !ext.Header.SystemComponent).OrderBy(ext => ext.Header.Name);
            return userExtensions.Select(e => new ExtensionInformation { Name = e.Header.Name, Identifier = e.Header.Identifier }).ToList();
        }

        public void InstallExtensions(IEnumerable<ExtensionInformation> extensions)
        {
            foreach (var extension in extensions)
            {
                var query = ExtensionRepository.CreateQuery<VSGalleryEntry>(false, true);
                query.ExecuteCompleted += query_ExecuteCompleted;
                query.SearchText = extension.Name;
                query.ExecuteAsync(extension);
            }
        }

        public void UnInstallExtensions(IEnumerable<ExtensionInformation> extensions)
        {
            var installedUserExtensions = ExtensionManager.GetInstalledExtensions().Where(e => !e.Header.SystemComponent);
            foreach (var extension in extensions)
            {
                var extensionInformation = extension;
                var userExtension = installedUserExtensions.Single(e => e.Header.Name == extensionInformation.Name && e.Header.Identifier == extensionInformation.Identifier);
                if (userExtension == null) continue;

                LogMessage(string.Format("Uninstalling {0}", userExtension.Header.Name));
                ExtensionManager.Uninstall(userExtension);
            }
        }

        void query_ExecuteCompleted(object sender, ExecuteCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                LogMessage(string.Format("Error while searching online for extension: {0}", e.Error.Message));
                return;
            }

            var extensionInformation = (ExtensionInformation)e.UserState;
            var entry = e.Results.Cast<VSGalleryEntry>().Single(r => r.Name == extensionInformation.Name && r.VsixID == extensionInformation.Identifier);
            if (entry == null)
            {
                LogMessage(string.Format("Could not find {0} in Online Repository", extensionInformation.Name));
                return;
            }
            ExtensionRepository.DownloadAsync(entry);
        }

        void ExtensionRepository_DownloadCompleted(object sender, DownloadCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                LogMessage(string.Format("Error while downloading extension: {0}", e.Error.Message));
                return;
            }

            var installableExtension = e.Payload;
            if (installableExtension == null)
                return;
            ExtensionManager.InstallAsync(installableExtension, false);
        }

        void ExtensionManager_InstallCompleted(object sender, InstallCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                LogMessage(string.Format("Error while installing extension: {0}", e.Error.Message));
                return;
            }
            LogMessage(string.Format("Installed {0}", e.Extension.Header.Name));
        }

        void LogMessage(string message)
        {
            if (Log != null)
                Log.Invoke(message);
        }

        IVsExtensionManager ExtensionManager { get; set; }
        IVsExtensionRepository ExtensionRepository { get; set; }
    }
}
