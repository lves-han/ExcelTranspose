using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Excel转置
{
    public partial class ExcelTransport : Form
    {
        private DialogResult dialogResult = DialogResult.Cancel;
        public ExcelTransport()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dialogResult = DialogResult.OK;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            RowAddress.Text = AreaSelection();
        }

        private string AreaSelection()
        {
            Application ExcelApp = Globals.ThisAddIn.Application;
            object selection = ExcelApp.InputBox("请选择", "请选择区域", Type: 8);
            if (selection is bool)
            {
                return "";
            }

            Range selectionRange = selection as Range;
            return selectionRange.Address;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            ColumnAddress.Text = AreaSelection();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void ExcelTransport_FormClosing(object sender, FormClosingEventArgs e)
        {
            //e.Cancel = true;
            this.DialogResult = dialogResult;
            //this.Close();
        }
    }
}
