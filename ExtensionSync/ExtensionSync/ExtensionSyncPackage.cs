using System;
using System.IO;
using System.Runtime.InteropServices;
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
    public sealed class ExtensionSyncPackage : Package, IVsShellPropertyEvents
    {
        protected override void Initialize()
        {
            base.Initialize();
            var shellService = GetService(typeof(SVsShell)) as IVsShell;
            if (shellService != null)
                ErrorHandler.ThrowOnFailure(shellService.AdviseShellPropertyChanges(this, out cookie));
        }

        protected override void Dispose(bool disposing)
        {
            //Hack: Really should be persisting the settings on DTE Shutdown, but could not
            //get the DTE events to fire as expected.
            PersistExtensionSettings();
            base.Dispose(disposing);
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

            var extensionsToInstall = persistedExtensionSettings.Except(installedUserExtensions).ToList();
            var extensionsToRemove = installedUserExtensions.Except(persistedExtensionSettings).ToList();

            extensionManager.InstallExtensions(extensionsToInstall);
            extensionManager.UnInstallExtensions(extensionsToRemove);
        }

        void OptionsPageSettingsDirectoryPathUpdated(string obj)
        {
            SynchronizeExtensions();
        }

        private void LogMessage(string message)
        {
            DebugPane.Activate();
            DebugPane.OutputString(string.Format("{0}: {1} \r\n", PackageName, message));
        }

        public int OnShellPropertyChange(int propid, object var)
        {
            if ((int)__VSSPROPID.VSSPROPID_Zombie == propid)
            {
                if ((bool)var == false)
                {
                    //Visual Studio is now ready and loaded up
                    var shellService = GetService(typeof(SVsShell)) as IVsShell;
                    if (shellService != null)
                        ErrorHandler.ThrowOnFailure(shellService.UnadviseShellPropertyChanges(cookie));
                    cookie = 0;
                    optionsPage = (OptionsPage)GetDialogPage(typeof(OptionsPage));
                    optionsPage.SettingsDirectoryPathUpdated += OptionsPageSettingsDirectoryPathUpdated;

                    var vsExtensionManager = GetService(typeof(SVsExtensionManager)) as IVsExtensionManager;
                    var vsExtensionRepository = GetService(typeof(SVsExtensionRepository)) as IVsExtensionRepository;

                    extensionManager = new ExtensionManagerFacade(vsExtensionManager, vsExtensionRepository);
                    extensionManager.Log += LogMessage;

                    SynchronizeExtensions();
                }
            }
            return VSConstants.S_OK;
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
                if (String.IsNullOrEmpty(optionsPage.SettingsDirectoryPath) || !Directory.Exists(optionsPage.SettingsDirectoryPath))
                {
                    LogMessage(string.Format("Invalid Directory configured for persisting settings. Defaulting to {0}",
                                             UserLocalDataPath));
                    optionsPage.SettingsDirectoryPath = UserLocalDataPath;
                }
                return Path.Combine(optionsPage.SettingsDirectoryPath, SettingsFileName);
            }
        }

        IVsOutputWindowPane debugPane;
        ExtensionManagerFacade extensionManager;
        OptionsPage optionsPage;
        uint cookie;

        const string PackageName = "ExtensionSync";
        const string SettingsFileName = "ExtensionSync.xml";
    }
}
