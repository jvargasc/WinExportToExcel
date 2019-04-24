using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WinExportToExcelCS
{
    public partial class Form1 : Form
    {

    #region Functions
        public Form1()
        {
            InitializeComponent();
        }

        private void TxtOnlyNumbers( object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if ((int)(e.KeyChar) != 8)
                if ((int)(e.KeyChar) < 48 || (int)(e.KeyChar) > 57)
                    e.Handled = true;
        }

        private void GetData()
        {
            DataTable TmpTableRead = new DataTable();

            //SqlConnection TmpConnection;
            using (SqlConnection TmpConnection = new SqlConnection(@"Data Source = bptjq4dxes.database.windows.net; Initial Catalog = Northwind; Persist Security Info = True; User ID = dbuser; Password = P4$$w0rd"))
            {

                TmpConnection.Open();

                StringBuilder CmdText = new StringBuilder();

                CmdText.Append("SELECT TOP " + TxtTopRows.Text.Trim() + "* \r\n");
                CmdText.Append("FROM [dbo].[Orders Qry]");

                SqlCommand TmpSqlCommand = new SqlCommand(CmdText.ToString(), TmpConnection);
                TmpSqlCommand.CommandType = CommandType.Text;

                SqlDataAdapter TmpSqlDataAdapter = new SqlDataAdapter();

                TmpSqlDataAdapter.SelectCommand = TmpSqlCommand;
                TmpSqlDataAdapter.Fill(TmpTableRead);

                if (TmpTableRead.Rows.Count > 0)
                    DtgOrders.DataSource = TmpTableRead;
                else
                {
                    MessageBox.Show("Data not found!!!");
                    DtgOrders.DataSource = null;
                }

            }

        }
        #endregion

        #region Controls
        private void TxtTopRows_KeyPress(object sender, KeyPressEventArgs e)
        {
            TxtOnlyNumbers(sender, e);
        }

        private void TxtTopRows_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                if (TxtTopRows.Text.Trim() != "")                
                    GetData();
            else
                MessageBox.Show("You must insert a numeric value");
        }

    #region Buttons
        private void BtnClear_Click(object sender, EventArgs e)
        {
            DtgOrders.DataSource = null;
        }

        private void BtnExcelMethod01_Click(object sender, EventArgs e)
        {
            if(DtgOrders.Rows.Count > 0)
            {

                string PathFile = Path.GetFullPath(Application.StartupPath);

                ExportToExcel TmpExportToExcel = new ExportToExcel
                {
                    ExcelWorksheetName = "Orders",
                    TmpGridView = DtgOrders,
                    TmpFileName = PathFile + @"\ExcelFile_" + DateTime.Now.ToString("yyyyMMddHHmmssffftt") + ".xlsx"
                };
                TmpExportToExcel.ExportData(1);

                TmpExportToExcel = null;
            }

        }

        private void BtnExcelMethod02_Click(object sender, EventArgs e)
        {
            if(DtgOrders.Rows.Count > 0)
            {
                string PathFile = Path.GetFullPath(Application.StartupPath);

                ExportToExcel TmpExportToExcel = new ExportToExcel
                {
                    ExcelWorksheetName = "Orders",
                    TmpDataTable = (DataTable)DtgOrders.DataSource,
                    TmpFileName = PathFile + @"\ExcelFile_" + DateTime.Now.ToString("yyyyMMddHHmmssffftt") + ".xlsx"
                };

                TmpExportToExcel.ExportData(2);

                TmpExportToExcel = null;

            }

        }

        #endregion

        #endregion

    }
}
