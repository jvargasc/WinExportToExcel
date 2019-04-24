Imports System.Text
Imports System.Data
Imports System.Data.SqlClient

Public Class FrmExportToExcel

#Region "Functions"
    Public Sub TxtOnlyNumbers(sender As Object, e As System.Windows.Forms.KeyPressEventArgs)

        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If (Asc(e.KeyChar) <> 8) Then
            If (Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57) Then
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub GetData()

        Dim TmpTableRead As DataTable = New DataTable()

        Using TmpConnection As SqlConnection = New SqlConnection("Data Source = bptjq4dxes.database.windows.net; Initial Catalog = Northwind; Persist Security Info = True; User ID = dbuser; Password = P4$$w0rd")

            TmpConnection.Open()

            Dim CmdText As StringBuilder = New StringBuilder()

            CmdText.Append("SELECT TOP " & TxtTopRows.Text.Trim() & " *" & vbNewLine)
            CmdText.Append("FROM [dbo].[Orders Qry]")

            Dim TmpSqlCommand As SqlCommand = New SqlCommand(CmdText.ToString(), TmpConnection)
            TmpSqlCommand.CommandType = CommandType.Text

            Dim TmpSqlDataAdapter As SqlDataAdapter = New SqlDataAdapter()

            TmpSqlDataAdapter.SelectCommand = TmpSqlCommand
            TmpSqlDataAdapter.Fill(TmpTableRead)

            If (TmpTableRead.Rows.Count > 0) Then
                DtgOrders.DataSource = TmpTableRead
            Else

                MessageBox.Show("Data not found!!!")
                DtgOrders.DataSource = Nothing
            End If

        End Using
    End Sub
#End Region

#Region "Controls"

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click
        DtgOrders.DataSource = Nothing
    End Sub

    Private Sub TxtTopRows_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtTopRows.KeyDown
        If (e.KeyCode = Keys.Enter) AndAlso (TxtTopRows.Text.Trim <> "") Then

            GetData()

        End If
    End Sub

    Private Sub TxtTopRows_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtTopRows.KeyPress

        TxtOnlyNumbers(sender, e)

    End Sub

    Private Sub BtnExcelMethod01_Click(sender As Object, e As EventArgs) Handles BtnExcelMethod01.Click

        If (DtgOrders.Rows.Count > 0) Then

            Dim PathFile As String = My.Application.Info.DirectoryPath

            Dim TmpExportToExcel As New ExportToExcel With {
                .ExcelWorksheetName = "Orders",
                .TmpGridView = DtgOrders,
                .TmpFileName = PathFile & "\ExcelFile_" & Now.ToString("yyyyMMddHHmmssffftt") & ".xlsx"
            }
            TmpExportToExcel.ExportData(1)

            TmpExportToExcel = Nothing

        End If

    End Sub

    Private Sub BtnExcelMethod02_Click(sender As Object, e As EventArgs) Handles BtnExcelMethod02.Click

        If (DtgOrders.Rows.Count > 0) Then

            Dim PathFile As String = My.Application.Info.DirectoryPath

            Dim TmpExportToExcel As New ExportToExcel With {
                .ExcelWorksheetName = "Orders",
                .TmpDataTable = DtgOrders.DataSource,
                .TmpFileName = PathFile & "\ExcelFile_" & Now.ToString("yyyyMMddHHmmssffftt") & ".xlsx"
            }
            TmpExportToExcel.ExportData(2)

            TmpExportToExcel = Nothing

        End If

    End Sub

#End Region

End Class
