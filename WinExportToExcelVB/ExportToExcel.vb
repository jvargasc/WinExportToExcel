
Imports ClosedXML.Excel

Public Class ExportToExcel

    Public TmpGridView As DataGridView
    Public TmpDataTable As DataTable
    Public TmpFileName As String
    Public ExcelWorksheetName As String

    Public Sub New()

        TmpGridView = New DataGridView

    End Sub

    Public Sub ExportData(ByVal ExcelOption As Byte)

        'If (TmpFileName.Trim <> "") And (TmpGridView.Rows.Count > 0) Then
        If (TmpFileName.Trim <> "") Then
            Select Case ExcelOption
                Case 1
                    GenerateFileMethod01()
                Case 2
                    GenerateFileMethod02()

            End Select
        Else
            MessageBox.Show("No hay data para generar.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub GenerateFileMethod01()

        Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application '= Microsoft.Office.Interop.Excel.Application();  
        Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = ExcelApp.Workbooks.Add(Type.Missing)
        Dim ExcelWorksheet As Microsoft.Office.Interop.Excel._Worksheet = Nothing

        ExcelWorksheet = ExcelWorkbook.Sheets("Sheet1")
        ExcelWorksheet = ExcelWorkbook.ActiveSheet

        ExcelWorksheet.Name = ExcelWorksheetName

        For i = 0 To (TmpGridView.Columns.Count - 1)
            ExcelWorksheet.Cells(1, i + 1) = TmpGridView.Columns(i).HeaderText
        Next

        For i = 0 To (TmpGridView.Rows.Count - 1)
            For j = 0 To (TmpGridView.Columns.Count - 1)
                If IsNothing(TmpGridView.Rows(i).Cells(j).Value) Then Exit For
                ExcelWorksheet.Cells(i + 2, j + 1) = TmpGridView.Rows(i).Cells(j).Value.ToString()
            Next

        Next

        ExcelWorksheet.Columns("B:B").Select
        ExcelApp.Selection.NumberFormat = "yyyy-mm-dd hh:mm:ss"
        ExcelApp.WindowState = ExcelApp.WindowState.xlMaximized
        ExcelWorksheet.Cells.Select()
        ExcelWorksheet.Cells.EntireColumn.AutoFit()
        ExcelWorksheet.Cells.Range("A1").Select()
        ExcelApp.Visible = True

        Try
            ExcelWorkbook.SaveAs(TmpFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing)
            MessageBox.Show("File successfully generated:" & vbNewLine & TmpFileName, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            ExcelApp = Nothing
        End Try

    End Sub

    Private Sub GenerateFileMethod02()

        Dim TmpXLWorkbook As XLWorkbook = New XLWorkbook()
        Dim p As New System.Diagnostics.Process
        Dim s As New System.Diagnostics.ProcessStartInfo(TmpFileName)

        Try
            TmpXLWorkbook.Worksheets.Add(TmpDataTable, "WorksheetName")
            TmpXLWorkbook.SaveAs(TmpFileName)

            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

            MessageBox.Show("File successfully generated:" & vbNewLine & vbNewLine & TmpFileName, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Finally
            TmpXLWorkbook = Nothing
            p = Nothing
            s = Nothing
        End Try

    End Sub

End Class
