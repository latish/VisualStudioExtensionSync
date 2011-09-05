using System;
using System.Windows.Forms;

namespace ExtensionSync
{
    public partial class OptionsUserControl : UserControl
    {
        public OptionsUserControl()
        {
            InitializeComponent();
        }

        public OptionsPage OptionsPage { get; set; }

        public void Initialize()
        {
            UpdatePathTextBoxes();
            cbxEnableLogging.Checked = OptionsPage.LoggingEnabled;
            btnBrowseLogFilePath.Enabled = cbxEnableLogging.Checked;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            UpdatePathTextBoxes();
        }

        private void UpdatePathTextBoxes()
        {
            txtSettingsFilePath.Text = OptionsPage.SettingsDirectoryPath;
            txtLogPath.Text = OptionsPage.LogDirectoryPath;
        }

        private void BtnBrowseSettingsPathClick(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK) return;

            txtSettingsFilePath.Text = folderBrowserDialog.SelectedPath;
            OptionsPage.SettingsDirectoryPath = txtSettingsFilePath.Text;
        }

        private void BtnBrowseLogFilePathClick(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK) return;

            txtLogPath.Text = folderBrowserDialog.SelectedPath;
            OptionsPage.LogDirectoryPath = txtLogPath.Text;
        }

        private void CbxEnableLoggingCheckedChanged(object sender, EventArgs e)
        {
            OptionsPage.LoggingEnabled = cbxEnableLogging.Checked;
            btnBrowseLogFilePath.Enabled = cbxEnableLogging.Checked;
        }

        private void CbxAutoUpdateExtensionsCheckedChanged(object sender, EventArgs e)
        {
            OptionsPage.AutoUpdateExtensions = cbxAutoUpdateExtensions.Checked;
        }
    }
}
