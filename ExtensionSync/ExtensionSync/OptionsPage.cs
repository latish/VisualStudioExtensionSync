using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using Microsoft.VisualStudio.Shell;
using System.Windows.Forms;

namespace LatishSehgal.ExtensionSync
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [CLSCompliant(false), ComVisible(true)]
    [Guid("5E6D92BF-3F5E-4C07-B965-B4EAA6D0D6A0")]
    public class OptionsPage : DialogPage
    {
        private OptionsUserControl optionsWindow;
        public event Action<string> SettingsDirectoryPathUpdated;

        [Category("Extension Sync"), DisplayName(@"Extensions File Path"), Description("Directory used to persist Extension Information")]
        public string SettingsDirectoryPath { get; set; }

        [Category("Extension Sync"), DisplayName(@"Log File Path"), Description("Directory used to store log file")]
        public string LogDirectoryPath { get; set; }

        [Category("Extension Sync"), DisplayName(@"Logging Enabled "), Description("Enable or disable logging")]
        public bool LoggingEnabled { get; set; }

        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        protected override IWin32Window Window
        {
            get
            {
                optionsWindow = new OptionsUserControl { OptionsPage = this };
                optionsWindow.Initialize();
                return optionsWindow;
            }
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            base.OnApply(e);
            OnSettingsDirectoryPathUpdated(SettingsDirectoryPath);
        }


        void OnSettingsDirectoryPathUpdated(string path)
        {
            if (SettingsDirectoryPathUpdated != null)
                SettingsDirectoryPathUpdated.Invoke(path);
        }
    }
}
