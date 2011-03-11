using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
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
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            var uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            var dte = GetService(typeof(SDTE)) as DTE;
            var extManager = GetService(typeof(SVsExtensionManager)) as IVsExtensionManager;
            var extRepository = GetService(typeof(SVsExtensionRepository)) as IVsExtensionRepository;
            var installedExtensions = extManager.GetInstalledExtensions();
            var userExtensions = installedExtensions.Where(ext => !ext.Header.SystemComponent);
            foreach (var extension in userExtensions)
            {
                Debug.WriteLine(String.Format("Name: {0}, Installed By MSI: {1}, Type: {2}, State: {3}, System Component: {4}", extension.Header.Name, extension.Header.InstalledByMsi, extension.Type,
                    extension.State, extension.Header.SystemComponent));
            }

            var entry = new VSGalleryEntry()
            {
                DownloadUrl = "http://visualstudiogallery.msdn.microsoft.com/1269c9a1-fcfe-4b47-91e7-22c7027f3c41/file/46303/1/KillCassini.vsix?SRC=VSIDE&amp;update=true",
                VsixID = "6ffccb42-5c12-4632-82d8-41d3349e8ba8",
                VsixReferences = string.Empty
            };
            var installableExtension = extRepository.Download(entry);
            var restartReason = extManager.Install(installableExtension, true);
        }

    }
}
