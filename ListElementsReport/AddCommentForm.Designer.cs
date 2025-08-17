namespace ListElementsReport
{
    partial class AddCommentForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        public System.ComponentModel.IContainer components = null;

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
        public void InitializeComponent()
        {
            this.comboBoxRangeValue = new System.Windows.Forms.ComboBox();
            this.listViewObjectsReport = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader10 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.buttonWriteObject = new System.Windows.Forms.Button();
            this.buttonWriteToReport = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.buttonSeparator = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonDeleteComment = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBoxRangeValue
            // 
            this.comboBoxRangeValue.FormattingEnabled = true;
            this.comboBoxRangeValue.Location = new System.Drawing.Point(27, 62);
            this.comboBoxRangeValue.Name = "comboBoxRangeValue";
            this.comboBoxRangeValue.Size = new System.Drawing.Size(179, 21);
            this.comboBoxRangeValue.TabIndex = 0;
            // 
            // listViewObjectsReport
            // 
            this.listViewObjectsReport.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewObjectsReport.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader10});
            this.listViewObjectsReport.FullRowSelect = true;
            this.listViewObjectsReport.GridLines = true;
            this.listViewObjectsReport.HideSelection = false;
            this.listViewObjectsReport.Location = new System.Drawing.Point(24, 12);
            this.listViewObjectsReport.Name = "listViewObjectsReport";
            this.listViewObjectsReport.Size = new System.Drawing.Size(723, 190);
            this.listViewObjectsReport.TabIndex = 13;
            this.listViewObjectsReport.UseCompatibleStateImageBehavior = false;
            this.listViewObjectsReport.View = System.Windows.Forms.View.Details;
            this.listViewObjectsReport.SelectedIndexChanged += new System.EventHandler(this.listViewObjectsReport_SelectedIndexChanged);
            this.listViewObjectsReport.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listViewObjectsReport_MouseDoubleClick);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Поз. обозначение";
            this.columnHeader1.Width = 110;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Наименование";
            this.columnHeader2.Width = 250;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Кол.";
            // 
            // columnHeader10
            // 
            this.columnHeader10.Text = "Примечание";
            this.columnHeader10.Width = 130;
            // 
            // buttonWriteObject
            // 
            this.buttonWriteObject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonWriteObject.Location = new System.Drawing.Point(232, 60);
            this.buttonWriteObject.Name = "buttonWriteObject";
            this.buttonWriteObject.Size = new System.Drawing.Size(466, 23);
            this.buttonWriteObject.TabIndex = 14;
            this.buttonWriteObject.Text = "Добавить в примечание";
            this.buttonWriteObject.UseVisualStyleBackColor = true;
            this.buttonWriteObject.Click += new System.EventHandler(this.buttonWriteObject_Click);
            // 
            // buttonWriteToReport
            // 
            this.buttonWriteToReport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonWriteToReport.Location = new System.Drawing.Point(24, 396);
            this.buttonWriteToReport.Name = "buttonWriteToReport";
            this.buttonWriteToReport.Size = new System.Drawing.Size(723, 23);
            this.buttonWriteToReport.TabIndex = 15;
            this.buttonWriteToReport.Text = "Записать в отчет";
            this.buttonWriteToReport.UseVisualStyleBackColor = true;
            this.buttonWriteToReport.Click += new System.EventHandler(this.buttonWriteToReport_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.buttonSeparator);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.comboBoxRangeValue);
            this.groupBox1.Controls.Add(this.buttonWriteObject);
            this.groupBox1.Location = new System.Drawing.Point(24, 231);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(723, 148);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Выбранный объект";
            // 
            // buttonSeparator
            // 
            this.buttonSeparator.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonSeparator.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.buttonSeparator.Location = new System.Drawing.Point(335, 89);
            this.buttonSeparator.Name = "buttonSeparator";
            this.buttonSeparator.Size = new System.Drawing.Size(30, 48);
            this.buttonSeparator.TabIndex = 19;
            this.buttonSeparator.Text = ";";
            this.buttonSeparator.UseVisualStyleBackColor = true;
            this.buttonSeparator.Click += new System.EventHandler(this.buttonSeparator_Click_1);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(19, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(313, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "Обязательным разделителем между номиналами является";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 15;
            // 
            // buttonDeleteComment
            // 
            this.buttonDeleteComment.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDeleteComment.Location = new System.Drawing.Point(359, 208);
            this.buttonDeleteComment.Name = "buttonDeleteComment";
            this.buttonDeleteComment.Size = new System.Drawing.Size(388, 23);
            this.buttonDeleteComment.TabIndex = 16;
            this.buttonDeleteComment.Text = "Удалить примечание";
            this.buttonDeleteComment.UseVisualStyleBackColor = true;
            this.buttonDeleteComment.Click += new System.EventHandler(this.buttonDeleteComment_Click);
            // 
            // AddCommentForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(766, 431);
            this.Controls.Add(this.buttonDeleteComment);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonWriteToReport);
            this.Controls.Add(this.listViewObjectsReport);
            this.MinimumSize = new System.Drawing.Size(642, 470);
            this.Name = "AddCommentForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавление примечаний для подборочных элементов";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SelectionRezCondForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SelectionRezCondForm_FormClosed);
            this.Load += new System.EventHandler(this.SelectionRezCondForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ComboBox comboBoxRangeValue;
        public System.Windows.Forms.ListView listViewObjectsReport;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader10;
        private System.Windows.Forms.Button buttonWriteObject;
        private System.Windows.Forms.Button buttonWriteToReport;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button buttonDeleteComment;
        private System.Windows.Forms.Button buttonSeparator;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label label1;
    }
}