using System;
using System.Windows.Forms;

namespace LatishSehgal.ExtensionSync
{
    public partial class OptionsUserControl : UserControl
    {
        public OptionsUserControl()
        {
            InitializeComponent();
        }

        internal OptionsPage optionsPage;

        public void Initialize()
        {
            txtPath.Text = optionsPage.DirectoryPath;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            txtPath.Text = optionsPage.DirectoryPath;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK) return;

            txtPath.Text = folderBrowserDialog.SelectedPath;
            UpdateDirectoryPath();
        }

        private void txtPath_TextChanged(object sender, EventArgs e)
        {
            UpdateDirectoryPath();
        }

        void UpdateDirectoryPath()
        {
            optionsPage.DirectoryPath = txtPath.Text;
        }
    }
}
