using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Excel转置
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelTransport transportForm = new ExcelTransport();
            DialogResult res = transportForm.ShowDialog();
            if (res == DialogResult.Cancel)
            {
                transportForm.Close();
                return;
            }

            string rowAddress = transportForm.RowAddress.Text;
            string columnAddress = transportForm.ColumnAddress.Text;
            if (rowAddress == "" || columnAddress == "")
            {
                return;
            }
            transportForm.Close();
            Application ExcelApp = Globals.ThisAddIn.Application;
            Worksheet sht = ExcelApp.ActiveSheet;
            Range selectionRow = sht.Range[rowAddress];
            Range selectionColumn = sht.Range[columnAddress];


            if (selectionRow.EntireRow.Address == rowAddress)
            {
                MessageBox.Show("请勿选择整行或整列","错误",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }


            
            object selection = ExcelApp.InputBox("请选择保存位置", "请选择", Type: 8);
            if (selection is bool)
            {
                return;
            }
            Range saveRange = selection as Range;
            saveRange = saveRange[1, 1];


            Range unionRange = ExcelApp.Intersect(selectionRow, selectionColumn);
            if (unionRange.Cells.Count <= 0)
            {
                MessageBox.Show("没有找到重合的区域，无法继续！", "选择错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            int beginRow = (unionRange[1, 1] as Range).Row + unionRange.Rows.Count;
            int beginColumn = (unionRange[1, 1] as Range).Column + unionRange.Columns.Count;

            int endRow = (selectionColumn[1, 1] as Range).Row + selectionColumn.Rows.Count;
            int endColumn = (selectionRow[1, 1] as Range).Column + selectionRow.Columns.Count;


            object[,] rowObjects = selectionRow.Value;
            object[,] columnObjects = selectionColumn.Value;

            int index = 0;
            List<object[]> result = new List<object[]>();
            for (int i = beginColumn; i < endColumn; i++)
            {
                for (int j = beginRow; j < endRow; j++)
                {
                    List<object> rangeKey = new List<object>();


                    for (int n = 1; n <= rowObjects.GetLength(0); n++)
                    {

                        rangeKey.Add(rowObjects[n, i]);
                    }
                    for (int k = 1; k <= columnObjects.GetLength(1); k++)
                    {
                        rangeKey.Add(columnObjects[j, k]);
                    }
                    object[] row = new object[rangeKey.Count + 1];
                    for (int k = 0; k < rangeKey.Count; k++)
                    {
                        row[ k] = rangeKey[k];
                    }
                    row[rangeKey.Count] = sht.Cells[j, i].Value;
                    result.Add(row);
                    //saveRange.Offset[index, 0].Resize[1, rangeKey.Count + 1].Value = row;
                    //saveRange.Offset[index, rangeKey.Count + 1].Resize[1, rangeKey.Count].Value = value;
                    index += 1;
                }
            }
            object[,] rows = new object[result.Count,result[0].Length];
            for (int i = 0; i < result.Count; i++)
            {
                object[] row = result[i];
                for (int j = 0; j < row.Length; j++)
                {
                    rows[i, j] = row[j];
                }
            }

            saveRange.Resize[rows.GetLength(0), rows.GetLength(1)].Value = rows;
            MessageBox.Show("完成");
        }
    }
}
