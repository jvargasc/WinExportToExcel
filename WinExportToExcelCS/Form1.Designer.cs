namespace WinExportToExcelCS
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.DtgOrders = new System.Windows.Forms.DataGridView();
            this.TxtTopRows = new System.Windows.Forms.TextBox();
            this.LblTopRows = new System.Windows.Forms.Label();
            this.BtnClear = new System.Windows.Forms.Button();
            this.BtnExcelMethod02 = new System.Windows.Forms.Button();
            this.BtnExcelMethod01 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DtgOrders)).BeginInit();
            this.SuspendLayout();
            // 
            // DtgOrders
            // 
            this.DtgOrders.AllowUserToAddRows = false;
            this.DtgOrders.AllowUserToDeleteRows = false;
            this.DtgOrders.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DtgOrders.Location = new System.Drawing.Point(12, 80);
            this.DtgOrders.Name = "DtgOrders";
            this.DtgOrders.ReadOnly = true;
            this.DtgOrders.Size = new System.Drawing.Size(930, 329);
            this.DtgOrders.TabIndex = 20;
            // 
            // TxtTopRows
            // 
            this.TxtTopRows.Location = new System.Drawing.Point(130, 36);
            this.TxtTopRows.Name = "TxtTopRows";
            this.TxtTopRows.Size = new System.Drawing.Size(126, 20);
            this.TxtTopRows.TabIndex = 0;
            this.TxtTopRows.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtTopRows_KeyDown);
            this.TxtTopRows.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtTopRows_KeyPress);
            // 
            // LblTopRows
            // 
            this.LblTopRows.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblTopRows.Location = new System.Drawing.Point(24, 34);
            this.LblTopRows.Name = "LblTopRows";
            this.LblTopRows.Size = new System.Drawing.Size(100, 23);
            this.LblTopRows.TabIndex = 8;
            this.LblTopRows.Text = "Top Rows";
            // 
            // BtnClear
            // 
            this.BtnClear.Location = new System.Drawing.Point(273, 34);
            this.BtnClear.Name = "BtnClear";
            this.BtnClear.Size = new System.Drawing.Size(75, 23);
            this.BtnClear.TabIndex = 1;
            this.BtnClear.Text = "Clear";
            this.BtnClear.UseVisualStyleBackColor = true;
            this.BtnClear.Click += new System.EventHandler(this.BtnClear_Click);
            // 
            // BtnExcelMethod02
            // 
            this.BtnExcelMethod02.Image = global::WinExportToExcelCS.Properties.Resources.Excel02;
            this.BtnExcelMethod02.Location = new System.Drawing.Point(854, 418);
            this.BtnExcelMethod02.Name = "BtnExcelMethod02";
            this.BtnExcelMethod02.Size = new System.Drawing.Size(88, 54);
            this.BtnExcelMethod02.TabIndex = 3;
            this.BtnExcelMethod02.UseVisualStyleBackColor = true;
            this.BtnExcelMethod02.Click += new System.EventHandler(this.BtnExcelMethod02_Click);
            // 
            // BtnExcelMethod01
            // 
            this.BtnExcelMethod01.Image = global::WinExportToExcelCS.Properties.Resources.Excel001;
            this.BtnExcelMethod01.Location = new System.Drawing.Point(757, 418);
            this.BtnExcelMethod01.Name = "BtnExcelMethod01";
            this.BtnExcelMethod01.Size = new System.Drawing.Size(88, 54);
            this.BtnExcelMethod01.TabIndex = 2;
            this.BtnExcelMethod01.UseVisualStyleBackColor = true;
            this.BtnExcelMethod01.Click += new System.EventHandler(this.BtnExcelMethod01_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(954, 484);
            this.Controls.Add(this.DtgOrders);
            this.Controls.Add(this.BtnExcelMethod02);
            this.Controls.Add(this.BtnExcelMethod01);
            this.Controls.Add(this.TxtTopRows);
            this.Controls.Add(this.LblTopRows);
            this.Controls.Add(this.BtnClear);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Export To Excel [CS]";
            ((System.ComponentModel.ISupportInitialize)(this.DtgOrders)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        internal System.Windows.Forms.DataGridView DtgOrders;
        internal System.Windows.Forms.Button BtnExcelMethod02;
        internal System.Windows.Forms.Button BtnExcelMethod01;
        internal System.Windows.Forms.TextBox TxtTopRows;
        internal System.Windows.Forms.Label LblTopRows;
        internal System.Windows.Forms.Button BtnClear;
    }
}

