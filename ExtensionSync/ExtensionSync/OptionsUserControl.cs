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

        public OptionsPage OptionsPage { get; set; }

        public void Initialize()
        {
            txtPath.Text = OptionsPage.DirectoryPath;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            txtPath.Text = OptionsPage.DirectoryPath;
        }

        private void BtnBrowseClick(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK) return;

            txtPath.Text = folderBrowserDialog.SelectedPath;
            UpdateDirectoryPath();
        }

        void UpdateDirectoryPath()
        {
            OptionsPage.DirectoryPath = txtPath.Text;
        }
    }
}
