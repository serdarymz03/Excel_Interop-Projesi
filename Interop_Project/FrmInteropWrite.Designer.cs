namespace Interop_Project
{
    partial class FrmInteropWrite
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
            this.DtgPersonList = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.TxtPath = new System.Windows.Forms.TextBox();
            this.BtnBrowse = new System.Windows.Forms.Button();
            this.BtnExport = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.DtgPersonList)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DtgPersonList
            // 
            this.DtgPersonList.AllowUserToAddRows = false;
            this.DtgPersonList.AllowUserToDeleteRows = false;
            this.DtgPersonList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DtgPersonList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DtgPersonList.Cursor = System.Windows.Forms.Cursors.PanWest;
            this.DtgPersonList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.DtgPersonList.Location = new System.Drawing.Point(3, 20);
            this.DtgPersonList.MultiSelect = false;
            this.DtgPersonList.Name = "DtgPersonList";
            this.DtgPersonList.ReadOnly = true;
            this.DtgPersonList.RowHeadersVisible = false;
            this.DtgPersonList.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DtgPersonList.Size = new System.Drawing.Size(896, 259);
            this.DtgPersonList.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.DtgPersonList);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(902, 282);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Liste   ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(28, 330);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 17);
            this.label1.TabIndex = 2;
            this.label1.Text = "Dosya Yolu : ";
            // 
            // TxtPath
            // 
            this.TxtPath.Enabled = false;
            this.TxtPath.Location = new System.Drawing.Point(129, 327);
            this.TxtPath.Name = "TxtPath";
            this.TxtPath.Size = new System.Drawing.Size(438, 24);
            this.TxtPath.TabIndex = 3;
            // 
            // BtnBrowse
            // 
            this.BtnBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.BtnBrowse.Location = new System.Drawing.Point(608, 320);
            this.BtnBrowse.Name = "BtnBrowse";
            this.BtnBrowse.Size = new System.Drawing.Size(88, 35);
            this.BtnBrowse.TabIndex = 4;
            this.BtnBrowse.Text = "Gözat";
            this.BtnBrowse.UseVisualStyleBackColor = true;
            this.BtnBrowse.Click += new System.EventHandler(this.BtnBrowse_Click);
            // 
            // BtnExport
            // 
            this.BtnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.BtnExport.Location = new System.Drawing.Point(727, 320);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(143, 35);
            this.BtnExport.TabIndex = 5;
            this.BtnExport.Text = "Tabloyu Çıkart";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // FrmInteropWrite
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(926, 379);
            this.Controls.Add(this.BtnExport);
            this.Controls.Add(this.BtnBrowse);
            this.Controls.Add(this.TxtPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FrmInteropWrite";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel Yazma";
            this.Load += new System.EventHandler(this.FrmInterop_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DtgPersonList)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView DtgPersonList;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TxtPath;
        private System.Windows.Forms.Button BtnBrowse;
        private System.Windows.Forms.Button BtnExport;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}