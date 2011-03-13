using System;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System.Linq;

namespace LatishSehgal.ExtensionSync
{
    [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidExtensionSyncPkgString)]
    [ProvideOptionPage(typeof(OptionsPage),
        "Extension Sync", "General", 0, 0, true)]
    public sealed class ExtensionSyncPackage : Package
    {
        protected override void Initialize()
        {
            base.Initialize();

            var vsExtensionManager = GetService(typeof(SVsExtensionManager)) as IVsExtensionManager;
            var vsExtensionRepository = GetService(typeof(SVsExtensionRepository)) as IVsExtensionRepository;

            extensionManager = new ExtensionManagerFacade(vsExtensionManager, vsExtensionRepository);
            extensionManager.Log += LogMessage;
            dte = ServiceProvider.GlobalProvider.GetService(typeof(SDTE)) as DTE;
            var dteEvents = dte.Events.DTEEvents;
            dteEvents.OnStartupComplete += OnStartupComplete;
        }

        protected override void Dispose(bool disposing)
        {
            PersistExtensionSettings();
            base.Dispose(disposing);
        }

        void OnStartupComplete()
        {
            SynchronizeExtensions();
        }

        void PersistExtensionSettings()
        {
            var settingsRepository = new SettingsRepository(extensionManager, SettingsFilePath);
            settingsRepository.Log += LogMessage;
            settingsRepository.PersistExtensionSettings();
        }

        void SynchronizeExtensions()
        {
            var settingsRepository = new SettingsRepository(extensionManager, SettingsFilePath);
            settingsRepository.Log += LogMessage;
            var persistedExtensionSettings = settingsRepository.GetPersistedExtensionSettings();

            if (persistedExtensionSettings.Count == 0)
                return;

            var installedUserExtensions = extensionManager.GetInstalledExtensionsInformation();

            var extensionsToInstall = persistedExtensionSettings.Except(installedUserExtensions);
            var extensionsToRemove = installedUserExtensions.Except(persistedExtensionSettings);

            extensionManager.InstallExtensions(extensionsToInstall);
            extensionManager.UnInstallExtensions(extensionsToRemove);
        }

        private void LogMessage(string message)
        {
            DebugPane.Activate();
            DebugPane.OutputString(string.Format("{0}: {1} \r\n", PackageName, message));
        }

        private IVsOutputWindowPane DebugPane
        {
            get
            {
                if (debugPane == null)
                {
                    var outputWindow = GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
                    var debugPaneGuid = VSConstants.GUID_OutWindowDebugPane;
                    if (outputWindow != null)
                        outputWindow.GetPane(ref debugPaneGuid, out debugPane);
                }
                return debugPane;
            }
        }

        string SettingsFilePath
        {
            get
            {
                var page = (OptionsPage)GetDialogPage(typeof(OptionsPage));
                var directoryPath = page.DirectoryPath;
                if (String.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath))
                {
                    LogMessage(string.Format("Invalid Directory configured for persisting settings. Defaulting to {0}",
                                             UserLocalDataPath));
                    directoryPath = page.DirectoryPath = UserLocalDataPath;
                }
                return Path.Combine(directoryPath, SettingsFileName);
            }
        }

        IVsOutputWindowPane debugPane;
        ExtensionManagerFacade extensionManager;
        DTE dte;

        const string PackageName = "ExtensionSync";
        const string SettingsFileName = "ExtensionSync.xml";
    }
}
