using Microsoft.Office.Interop.Excel;
using Nilsen.Framework.Factory.Objects.Classes.Services;
using System;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace Races_CSV_To_Excel
{
    public partial class frmRacesCSVtoExcel : Form
    {
        private String _SavePath = string.Format("{0}\\Flicker City Productions\\RacesCSVToExcel\\files", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));

        public frmRacesCSVtoExcel()
        {
            InitializeComponent();
        }
        private void RacesCSVtoExcel_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btnFileBrowse_Click(object sender, EventArgs e)
        {
            ofdRacesCSV.ShowDialog();
        }
        private void ofdRacesCSV_FileOk(object sender, CancelEventArgs e)
        {
            txtFileName.Text = ofdRacesCSV.FileName;
        }
        protected override bool ProcessDialogKey(Keys keyData)
        {
            int keyCode = (int)keyData;

            return base.ProcessDialogKey(keyData);
        }
        private void btnProcess_Click(object sender, EventArgs e)
        {
            FileInfo fi;
            fi = new FileInfo(ofdRacesCSV.FileName);

            if (fi.Exists)
            {
                if (fi.Extension.ToLower().Equals(".csv"))
                {
                    var service = ServiceFactory.GetReportService(txtConsole, btnProcess, ddReports.SelectedIndex);
                    Thread reportThread1 = new Thread(() => service.CreateExcelFile(fi));
                    reportThread1.Start();
                }
                else
                {
                    MessageBox.Show("File is not in the correct format.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("File Does Not Exist", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
        private void UpdateConsoleText(String sText, bool clear)
        {
            txtConsole.Text = clear ? string.Empty : txtConsole.Text;

            if (txtConsole.Text.Equals(string.Empty))
            {
                txtConsole.Text = string.Format("{0} >> {1}", txtConsole.Text, sText);
            }
            else
            {
                txtConsole.Text = string.Format("{0}\r\n\r\n >> {1}", txtConsole.Text, sText);
            }
            txtConsole.Refresh();
        }
    }
}
