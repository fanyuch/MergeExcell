namespace MergeExcel
{
    partial class ExcelMerge
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.cbxLeft = new System.Windows.Forms.ComboBox();
            this.cbxRight = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnEmrge = new System.Windows.Forms.Button();
            this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // cbxLeft
            // 
            this.cbxLeft.FormattingEnabled = true;
            this.cbxLeft.Location = new System.Drawing.Point(85, 12);
            this.cbxLeft.Name = "cbxLeft";
            this.cbxLeft.Size = new System.Drawing.Size(246, 20);
            this.cbxLeft.TabIndex = 0;
            this.cbxLeft.SelectedIndexChanged += new System.EventHandler(this.cbxLeft_SelectedIndexChanged);
            // 
            // cbxRight
            // 
            this.cbxRight.FormattingEnabled = true;
            this.cbxRight.Location = new System.Drawing.Point(498, 12);
            this.cbxRight.Name = "cbxRight";
            this.cbxRight.Size = new System.Drawing.Size(246, 20);
            this.cbxRight.TabIndex = 1;
            this.cbxRight.SelectedIndexChanged += new System.EventHandler(this.cbxRight_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "选择源文件";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(420, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "选择源文件";
            // 
            // btnEmrge
            // 
            this.btnEmrge.Location = new System.Drawing.Point(846, 12);
            this.btnEmrge.Name = "btnEmrge";
            this.btnEmrge.Size = new System.Drawing.Size(63, 20);
            this.btnEmrge.TabIndex = 4;
            this.btnEmrge.Text = "合并";
            this.btnEmrge.UseVisualStyleBackColor = true;
            this.btnEmrge.Click += new System.EventHandler(this.btnEmrge_Click);
            // 
            // ultraGrid1
            // 
            this.ultraGrid1.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free;
            this.ultraGrid1.DisplayLayout.Override.ColumnSizingArea = Infragistics.Win.UltraWinGrid.ColumnSizingArea.CellsOnly;
            this.ultraGrid1.DisplayLayout.Override.RowSelectorNumberStyle = Infragistics.Win.UltraWinGrid.RowSelectorNumberStyle.VisibleIndex;
            this.ultraGrid1.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFixed;
            this.ultraGrid1.DisplayLayout.UseFixedHeaders = true;
            this.ultraGrid1.Location = new System.Drawing.Point(9, 51);
            this.ultraGrid1.Name = "ultraGrid1";
            this.ultraGrid1.Size = new System.Drawing.Size(989, 503);
            this.ultraGrid1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(344, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(63, 20);
            this.button1.TabIndex = 7;
            this.button1.Text = "选择文件";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(757, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(63, 20);
            this.button2.TabIndex = 8;
            this.button2.Text = "选择文件";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(935, 12);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(63, 20);
            this.button4.TabIndex = 10;
            this.button4.Text = "导出";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // ExcelMerge
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1011, 569);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.ultraGrid1);
            this.Controls.Add(this.btnEmrge);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbxRight);
            this.Controls.Add(this.cbxLeft);
            this.MaximizeBox = false;
            this.Name = "ExcelMerge";
            this.Text = "发货单合并工具";
            ((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbxLeft;
        private System.Windows.Forms.ComboBox cbxRight;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnEmrge;
        private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button4;
    }
}

