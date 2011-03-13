using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using System.Xml.Serialization;
using EnvDTE;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.ExtensionManager.UI;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.Linq;

namespace LatishSehgal.ExtensionSync
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidExtensionSyncPkgString)]
    [ProvideOptionPage(typeof(OptionsPage),
        "Extension Sync", "General", 0, 0, true)]
    public sealed class ExtensionSyncPackage : Package
    {
        public ExtensionSyncPackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this));
        }

        protected override void Initialize()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this));
            base.Initialize();

            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                var menuCommandID = new CommandID(GuidList.guidExtensionSyncCmdSet, (int)PkgCmdIDList.cmdidSyncExtensions);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }

            var vsExtensionManager =  GetService(typeof (SVsExtensionManager)) as IVsExtensionManager;
            var vsExtensionRepository = GetService(typeof (SVsExtensionRepository)) as IVsExtensionRepository;

            extensionManager = new ExtensionManagerFacade(vsExtensionManager, vsExtensionRepository);
            extensionManager.Log += LogMessage;
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            var uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            var dte = GetService(typeof(SDTE)) as DTE;
            SynchronizeExtensions();
            //PersistExtensionSettings();
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
            var persistedExtensionSettings =settingsRepository.GetPersistedExtensionSettings();

            var installedUserExtensions = extensionManager.GetInstalledExtensionsInformation();

            var extensionsToInstall = persistedExtensionSettings.Except(installedUserExtensions);
            var extensionsToRemove = installedUserExtensions.Except(persistedExtensionSettings);

            extensionManager.InstallExtensions(extensionsToInstall);
            extensionManager.UnInstallExtensions(extensionsToRemove);
        }

        private void LogMessage(string message)
        {
            DebugPane.Activate();
            DebugPane.OutputString(string.Format("{0}: {1} \r\n",PackageName,message));
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

        const string PackageName = "ExtensionSync";
        const string SettingsFileName = "ExtensionSync.xml";
    }
}
