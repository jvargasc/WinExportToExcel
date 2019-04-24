<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmExportToExcel
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.BtnClear = New System.Windows.Forms.Button()
        Me.LblTopRows = New System.Windows.Forms.Label()
        Me.TxtTopRows = New System.Windows.Forms.TextBox()
        Me.DtgOrders = New System.Windows.Forms.DataGridView()
        Me.BtnExcelMethod02 = New System.Windows.Forms.Button()
        Me.BtnExcelMethod01 = New System.Windows.Forms.Button()
        CType(Me.DtgOrders, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnClear
        '
        Me.BtnClear.Location = New System.Drawing.Point(273, 34)
        Me.BtnClear.Name = "BtnClear"
        Me.BtnClear.Size = New System.Drawing.Size(75, 23)
        Me.BtnClear.TabIndex = 1
        Me.BtnClear.Text = "Clear"
        Me.BtnClear.UseVisualStyleBackColor = True
        '
        'LblTopRows
        '
        Me.LblTopRows.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTopRows.Location = New System.Drawing.Point(24, 34)
        Me.LblTopRows.Name = "LblTopRows"
        Me.LblTopRows.Size = New System.Drawing.Size(100, 23)
        Me.LblTopRows.TabIndex = 1
        Me.LblTopRows.Text = "Top Rows"
        '
        'TxtTopRows
        '
        Me.TxtTopRows.Location = New System.Drawing.Point(130, 36)
        Me.TxtTopRows.Name = "TxtTopRows"
        Me.TxtTopRows.Size = New System.Drawing.Size(126, 20)
        Me.TxtTopRows.TabIndex = 0
        '
        'DtgOrders
        '
        Me.DtgOrders.AllowUserToAddRows = False
        Me.DtgOrders.AllowUserToDeleteRows = False
        Me.DtgOrders.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DtgOrders.Location = New System.Drawing.Point(12, 80)
        Me.DtgOrders.Name = "DtgOrders"
        Me.DtgOrders.ReadOnly = True
        Me.DtgOrders.Size = New System.Drawing.Size(930, 329)
        Me.DtgOrders.TabIndex = 5
        '
        'BtnExcelMethod02
        '
        Me.BtnExcelMethod02.Image = Global.WinExportToExcelVB.My.Resources.Resources.Excel02
        Me.BtnExcelMethod02.Location = New System.Drawing.Point(854, 418)
        Me.BtnExcelMethod02.Name = "BtnExcelMethod02"
        Me.BtnExcelMethod02.Size = New System.Drawing.Size(88, 54)
        Me.BtnExcelMethod02.TabIndex = 3
        Me.BtnExcelMethod02.UseVisualStyleBackColor = True
        '
        'BtnExcelMethod01
        '
        Me.BtnExcelMethod01.Image = Global.WinExportToExcelVB.My.Resources.Resources.Excel001
        Me.BtnExcelMethod01.Location = New System.Drawing.Point(757, 418)
        Me.BtnExcelMethod01.Name = "BtnExcelMethod01"
        Me.BtnExcelMethod01.Size = New System.Drawing.Size(88, 54)
        Me.BtnExcelMethod01.TabIndex = 2
        Me.BtnExcelMethod01.UseVisualStyleBackColor = True
        '
        'FrmExportToExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(954, 484)
        Me.Controls.Add(Me.DtgOrders)
        Me.Controls.Add(Me.BtnExcelMethod02)
        Me.Controls.Add(Me.BtnExcelMethod01)
        Me.Controls.Add(Me.TxtTopRows)
        Me.Controls.Add(Me.LblTopRows)
        Me.Controls.Add(Me.BtnClear)
        Me.MaximizeBox = False
        Me.Name = "FrmExportToExcel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Export To Excel [VB]"
        CType(Me.DtgOrders, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BtnClear As Button
    Friend WithEvents LblTopRows As Label
    Friend WithEvents TxtTopRows As TextBox
    Friend WithEvents BtnExcelMethod01 As Button
    Friend WithEvents BtnExcelMethod02 As Button
    Friend WithEvents DtgOrders As DataGridView
End Class
