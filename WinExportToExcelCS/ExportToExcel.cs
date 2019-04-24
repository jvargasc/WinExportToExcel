using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;


namespace WinExportToExcelCS
{
    public class ExportToExcel
    {

        public DataGridView TmpGridView;
        public DataTable TmpDataTable;
        public String TmpFileName;
        public String ExcelWorksheetName;

        public ExportToExcel()
        {
            TmpGridView = new DataGridView();
        }

        public void ExportData(Byte ExcelOption)
        {

            if (TmpFileName.Trim() != "")
                switch (ExcelOption)
                {
                    case 1:
                        GenerateFileMethod01();
                        break;
                    case 2:
                        GenerateFileMethod02();
                        break;
                }
            else
                MessageBox.Show("No hay data para generar.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
        
        }

        public void GenerateFileMethod01()
        {

            Microsoft.Office.Interop.Excel.Application ExcelApp = new  Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook = ExcelApp.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet ExcelWorksheet = null;

            ExcelWorksheet = ExcelWorkbook.Sheets["Sheet1"];
            ExcelWorksheet = ExcelWorkbook.ActiveSheet;

            ExcelWorksheet.Name = ExcelWorksheetName;

            for (int i = 0; i < TmpGridView.Columns.Count - 1; i++)
                ExcelWorksheet.Cells[1, i + 1] = TmpGridView.Columns[i].HeaderText;
            
            for(int i = 0; i < TmpGridView.Rows.Count - 1; i++)
                for(int j = 0; j < TmpGridView.Columns.Count - 1; j++)
                {
                    if (TmpGridView.Rows[i].Cells[j].Value == null)
                        break;
                    ExcelWorksheet.Cells[i + 2, j + 1] = TmpGridView.Rows[i].Cells[j].Value.ToString();
                }

            ExcelWorksheet.Columns["B:B"].Select();
            ExcelApp.Selection.NumberFormat = "yyyy-mm-dd hh:mm:ss";
            ExcelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
            ExcelWorksheet.Cells.Select();
            ExcelWorksheet.Cells.EntireColumn.AutoFit();
            ExcelWorksheet.Cells.Range["A1"].Select();
            ExcelApp.Visible = true;

            try
            {
                ExcelWorkbook.SaveAs(TmpFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MessageBox.Show("File successfully generated:" + "\r\n" + TmpFileName, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                ExcelApp = null;
            }

        }

        private void GenerateFileMethod02()
        {

            XLWorkbook TmpXLWorkbook = new XLWorkbook();
            Process p = new Process();
            ProcessStartInfo s = new ProcessStartInfo(TmpFileName);

            try
            {
                TmpXLWorkbook.Worksheets.Add(TmpDataTable, "WorksheetName");
                TmpXLWorkbook.SaveAs(TmpFileName );

                s.UseShellExecute = true;
                s.WindowStyle = ProcessWindowStyle.Normal;
                p.StartInfo = s;
                p.Start();

                MessageBox.Show("File successfully generated:" + "\r\n\r\n" + TmpFileName, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                TmpXLWorkbook = null;
                p = null;
                s = null;
            }

        }

    }
}
