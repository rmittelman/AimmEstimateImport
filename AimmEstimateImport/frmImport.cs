using System;
using Aimm.Logging;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Configuration;
using WindowsInstaller;

namespace AimmEstimateImport
{

    public partial class frmImport : Form
    {
        clsImport imp;
        public frmImport()
        {
            InitializeComponent();
            imp = new clsImport();
            imp.StatusChanged += imp_StatusChanged;
            bool isIDE = (Debugger.IsAttached == true);
            string settingsPath;
            if(isIDE)
                settingsPath = Path.GetDirectoryName(Application.ExecutablePath);
            else
                settingsPath = Application.CommonAppDataPath.Remove(Application.CommonAppDataPath.LastIndexOf("."));
            imp.InitClass(settingsPath);
            string msiVersion = imp.MsiVersion;

            // Check for new version
            // Requires GetMSIVersion
            string myVersion = Application.ProductVersion.Substring(0, Application.ProductVersion.LastIndexOf("."));
            bool newerVersionAvailable = myVersion.CompareTo(msiVersion) < 0;
            if(newerVersionAvailable)
            {
                string msg = $"You have version { myVersion}, but version { imp.MsiVersion} is available.\n\nDo you want to upgrade now?";
                if(MessageBox.Show(msg, "New Version", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                {
                    Process.Start(imp.SourcePath);
                    imp = null;
                    Environment.Exit(0);
                }
            }
            this.lblVersion.Text = $"v {myVersion}";
        }

        ~frmImport()
        {
            imp = null;
        }

        #region objects

        ToolTip toolTip1 = new ToolTip();

        #endregion

        #region properties

        private string _status;
        public string Status
        {
            set
            {
                _status = value;
                txtStatus.Text = value;
            }
            get { return _status; }
        }

        #endregion

        #region events

        private void imp_StatusChanged(object sender, StatusChangedEventArgs e)
        {
            Status = e.Status;
        }

        private void frmImport_Load(object sender, EventArgs e)
        {
            LogIt.LogMethod();

            // get form position and size, and apply
            FormWindowState state = Properties.Settings.Default.wState;
            Point location = Properties.Settings.Default.wLocation;
            Size size = Properties.Settings.Default.wSize;
            WindowState = state == FormWindowState.Minimized ? FormWindowState.Normal : state;
            Location = location == new Point(0, 0) ? new Point(100, 100) : location;
            Size = size == new Size(0, 0) ? new Size(1030, 400) : size;

            // set tooltips
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            toolTip1.SetToolTip(btnFindExcel, "Find Excel ECM File");
            toolTip1.SetToolTip(btnImport, "Import AIMM Estimate from Excel ECM File");
        }

        private void frmImport_FormClosing(object sender, FormClosingEventArgs e)
        {
            // save current position and normal size
            Properties.Settings.Default.wState = WindowState;
            Properties.Settings.Default.wLocation = WindowState == FormWindowState.Normal ? Location : RestoreBounds.Location;
            Properties.Settings.Default.wSize = WindowState == FormWindowState.Normal ? Size : RestoreBounds.Size;
            Properties.Settings.Default.Save();
        }

        private void btnFindExcel_Click(object sender, EventArgs e)
        {
            using(OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.InitialDirectory = imp.SourcePath;
                ofd.Filter = "Excel files (*.xlsx, *.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*";
                ofd.FilterIndex = 1;
                if(ofd.ShowDialog() == DialogResult.OK)
                {
                    txtExcelFile.Text = ofd.FileName;
                    imp.ExcelFile = ofd.FileName;
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            btnImport.Enabled = false;
            imp.ImportECM();
        }

        private void txtExcelFile_TextChanged(object sender, EventArgs e)
        {
            btnImport.Enabled = File.Exists(txtExcelFile.Text);
        }

        #endregion
    }
}
