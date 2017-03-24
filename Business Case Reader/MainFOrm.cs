using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace Business_Case_Reader
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            btn_browseBC.Enabled = true;
            bth_ReadSheet.Enabled = false;
            btn_ExportCSV.Enabled = false;
        }

        private void btn_browseBC_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                openExcelFile(openFileDialog.FileName);
                
            }
        }

        public void openExcelFile(string filePath)
        {
            
            pic_Loading.Visible = true;
            tb_FileName.Text = Path.GetFileName(filePath);
            Excel.Application oXL = new Excel.Application();
            Excel.Workbook oWB = oXL.Workbooks.Open(filePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //read Excel sheets 
            foreach (Excel.Worksheet ws in oWB.Sheets)
            {
                lb_SheetNames.Items.Add(ws.Name);
            }
           
            //oWB.Close(false, Missing.Value, Missing.Value);
            pic_Loading.Visible = false;

            btn_browseBC.Enabled = false;
            bth_ReadSheet.Enabled = true;
            btn_ExportCSV.Enabled = false;

            tabovi.SelectedIndex = 1;
        }

    }

    public class FileNames
    {
        public string FileName { get; set; }
    }
}
