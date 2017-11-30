using System;
using Aimm.Logging;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

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
                settingsPath = Path.GetDirectoryName(Application.CommonAppDataPath);
            imp.InitClass(settingsPath);
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

            // set tooltips
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            toolTip1.SetToolTip(btnFindExcel, "Find Excel ECM File");
            toolTip1.SetToolTip(btnImport, "Import AIMM Estimate from Excel ECM File");
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
            imp.ImportExcel();
        }

        #endregion
    }
}
