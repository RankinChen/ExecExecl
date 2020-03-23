using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExecExecl
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnBrowseSourceFile_Click(object sender, EventArgs e)
        {
            DialogResult result = ofdBrowserFile.ShowDialog();
            if (result == DialogResult.OK)
            {
                var sourceFilePath = ofdBrowserFile.FileName;
                var sourceFileName = ofdBrowserFile.SafeFileName;
                txtSourceFile.Text = sourceFilePath;
            }
        }

        private void btnExec_Click(object sender, EventArgs e)
        {
            string newFilePath = ExecExcelFile.ReadAndGenerateExcel(txtSourceFile.Text.Trim());

            FileInfo fileInfo = new FileInfo(newFilePath);
            sfdSaveNewFile.FileName = fileInfo.Name;
            DialogResult result = sfdSaveNewFile.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileInfo.CopyTo(sfdSaveNewFile.FileName, true);
                try
                {
                    Process p = new Process();
                    p.StartInfo.FileName = "explorer.exe";
                    p.StartInfo.Arguments = $@" /select, {sfdSaveNewFile.FileName}";
                    p.Start();
                }
                finally
                {

                }
            }
        }
    }
}
