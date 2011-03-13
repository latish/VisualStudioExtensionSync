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

            ExtensionRepository.DownloadCompleted += ExtensionRepository_DownloadCompleted;
            ExtensionManager.InstallCompleted += ExtensionManager_InstallCompleted;
        }



        private void MenuItemCallback(object sender, EventArgs e)
        {
            var uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            var dte = GetService(typeof(SDTE)) as DTE;
            SynchronizeExtensions();
            //PersistExtensionSettings();
        }

        void SynchronizeExtensions()
        {
            var persistedExtensionSettings = GetPersistedExtensionSettings();
            var installedUserExtensions = GetInstalledExtensionsInformation();

            var extensionsToInstall = persistedExtensionSettings.Except(installedUserExtensions);
            var extensionsToRemove = installedUserExtensions.Except(persistedExtensionSettings);


            DownloadAndInstallExtensions(extensionsToInstall);
        }

        void DownloadAndInstallExtensions(IEnumerable<ExtensionInformation> extensionsToInstall)
        {
            var query = ExtensionRepository.CreateQuery<VSGalleryEntry>(false, true);
            query.ExecuteCompleted += query_ExecuteCompleted;

            foreach (var extensionInformation in extensionsToInstall)
            {
                query.SearchText = extensionInformation.Name;
                query.ExecuteAsync(extensionInformation);
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

            LogMessage(string.Format("Installed extension: {0}", e.Extension.Header.Name));

        }

        void PersistExtensionSettings()
        {
            var installedExtensionsInformation = GetInstalledExtensionsInformation();

            using (var fileStream = new FileStream(SettingsFilePath, FileMode.OpenOrCreate))
            {
                var serializer = new XmlSerializer(typeof(List<ExtensionInformation>));
                serializer.Serialize(fileStream, installedExtensionsInformation);
            }
        }

        List<ExtensionInformation> GetPersistedExtensionSettings()
        {
            var settings = new List<ExtensionInformation>();
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    using (var fileStream = new FileStream(SettingsFilePath, FileMode.Open))
                    {
                        var serializer = new XmlSerializer(typeof(List<ExtensionInformation>));
                        settings = (List<ExtensionInformation>)serializer.Deserialize(fileStream);
                        return settings;
                    }
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }

            return settings;
        }

        List<ExtensionInformation> GetInstalledExtensionsInformation()
        {
            var installedExtensions = ExtensionManager.GetInstalledExtensions();
            var userExtensions = installedExtensions.Where(ext => !ext.Header.SystemComponent);
            return userExtensions.Select(e => new ExtensionInformation { Name = e.Header.Name, Identifier = e.Header.Identifier }).ToList();
        }

        private void LogMessage(string message)
        {
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

        IVsExtensionManager ExtensionManager
        {
            get { return extensionManager ?? (extensionManager = GetService(typeof(SVsExtensionManager)) as IVsExtensionManager); }
        }

        IVsExtensionRepository ExtensionRepository
        {
            get { return extensionRepository ?? (extensionRepository = GetService(typeof(SVsExtensionRepository)) as IVsExtensionRepository); }
        }

        string SettingsFilePath
        {
            get
            {
                var page = (OptionsPage)GetDialogPage(typeof(OptionsPage));
                var directoryPath = page.DirectoryPath;
                if (String.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath))
                {
                    directoryPath = page.DirectoryPath = UserLocalDataPath;
                }
                return Path.Combine(directoryPath, SettingsFileName);
            }
        }

        IVsOutputWindowPane debugPane;
        IVsExtensionManager extensionManager;
        IVsExtensionRepository extensionRepository;

        const string PackageName = "ExtensionSync";
        const string SettingsFileName = "ExtensionSync.xml";
    }
}
