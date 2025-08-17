namespace ListElementsReport
{
    partial class AddDescriptionForm
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
            this.buttonDeleteName = new System.Windows.Forms.Button();
            this.listViewObjectsReport = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBoxName = new System.Windows.Forms.ComboBox();
            this.buttonWriteObject = new System.Windows.Forms.Button();
            this.buttonWriteToReport = new System.Windows.Forms.Button();
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonDeleteName
            // 
            this.buttonDeleteName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDeleteName.Enabled = false;
            this.buttonDeleteName.Location = new System.Drawing.Point(265, 241);
            this.buttonDeleteName.Name = "buttonDeleteName";
            this.buttonDeleteName.Size = new System.Drawing.Size(446, 23);
            this.buttonDeleteName.TabIndex = 25;
            this.buttonDeleteName.Text = "Удалить наименование";
            this.buttonDeleteName.UseVisualStyleBackColor = true;
            this.buttonDeleteName.Click += new System.EventHandler(this.buttonDeleteName_Click);
            // 
            // listViewObjectsReport
            // 
            this.listViewObjectsReport.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewObjectsReport.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader4,
            this.columnHeader5,
            this.columnHeader3});
            this.listViewObjectsReport.FullRowSelect = true;
            this.listViewObjectsReport.GridLines = true;
            this.listViewObjectsReport.HideSelection = false;
            this.listViewObjectsReport.Location = new System.Drawing.Point(12, 12);
            this.listViewObjectsReport.Name = "listViewObjectsReport";
            this.listViewObjectsReport.Size = new System.Drawing.Size(699, 221);
            this.listViewObjectsReport.TabIndex = 22;
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
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.comboBoxName);
            this.groupBox1.Controls.Add(this.buttonWriteObject);
            this.groupBox1.Location = new System.Drawing.Point(12, 266);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(699, 117);
            this.groupBox1.TabIndex = 24;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Выбранный объект";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 15;
            // 
            // comboBoxName
            // 
            this.comboBoxName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxName.FormattingEnabled = true;
            this.comboBoxName.Location = new System.Drawing.Point(6, 62);
            this.comboBoxName.Name = "comboBoxName";
            this.comboBoxName.Size = new System.Drawing.Size(687, 21);
            this.comboBoxName.TabIndex = 0;
            // 
            // buttonWriteObject
            // 
            this.buttonWriteObject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonWriteObject.Location = new System.Drawing.Point(307, 89);
            this.buttonWriteObject.Name = "buttonWriteObject";
            this.buttonWriteObject.Size = new System.Drawing.Size(386, 23);
            this.buttonWriteObject.TabIndex = 14;
            this.buttonWriteObject.Text = "Заменить наименование";
            this.buttonWriteObject.UseVisualStyleBackColor = true;
            this.buttonWriteObject.Click += new System.EventHandler(this.buttonWriteObject_Click);
            // 
            // buttonWriteToReport
            // 
            this.buttonWriteToReport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonWriteToReport.Location = new System.Drawing.Point(12, 399);
            this.buttonWriteToReport.Name = "buttonWriteToReport";
            this.buttonWriteToReport.Size = new System.Drawing.Size(699, 23);
            this.buttonWriteToReport.TabIndex = 23;
            this.buttonWriteToReport.Text = "Записать в отчет";
            this.buttonWriteToReport.UseVisualStyleBackColor = true;
            this.buttonWriteToReport.Click += new System.EventHandler(this.buttonWriteToReport_Click);
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Наименование ус-ва";
            this.columnHeader4.Width = 120;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "Обозначение уст-ва";
            this.columnHeader5.Width = 115;
            // 
            // AddDescriptionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(723, 434);
            this.Controls.Add(this.buttonDeleteName);
            this.Controls.Add(this.listViewObjectsReport);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonWriteToReport);
            this.MinimumSize = new System.Drawing.Size(498, 473);
            this.Name = "AddDescriptionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавление наименования";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.AddDescriptionForm_FormClosed);
            this.Load += new System.EventHandler(this.AddDescriptionForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonDeleteName;
        public System.Windows.Forms.ListView listViewObjectsReport;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.ComboBox comboBoxName;
        private System.Windows.Forms.Button buttonWriteObject;
        private System.Windows.Forms.Button buttonWriteToReport;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader5;
    }
}