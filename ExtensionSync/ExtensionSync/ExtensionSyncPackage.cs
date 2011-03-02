using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.VisualStudio.ExtensionManager;
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
            Trace.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this));
            base.Initialize();

            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                var menuCommandID = new CommandID(GuidList.guidExtensionSyncCmdSet, (int)PkgCmdIDList.cmdidSyncExtensions);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            var uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            var dte = GetService(typeof(SDTE)) as DTE;
            dynamic extManager = GetService(typeof(SVsExtensionManager)) ;
            IEnumerable<IInstalledExtension > installedExtensions = extManager.GetInstalledExtensions();
            foreach (var installedExtension in installedExtensions)
            {
                Debug.WriteLine(String.Format("Name: {0}, Installed By MSI: {1}, Type: {2}, State: {3}", installedExtension.Header.Name, installedExtension.Header.InstalledByMsi, installedExtension.Type, installedExtension.State));
            }
        }

    }
}
