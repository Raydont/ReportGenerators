namespace ListElementsReport
{
    partial class AddAdmissionForm
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
            this.listViewObjectsReport = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.comboBoxAdmissions = new System.Windows.Forms.ComboBox();
            this.buttonWriteObject = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonWriteToReport = new System.Windows.Forms.Button();
            this.buttonDeleteAdmission = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listViewObjectsReport
            // 
            this.listViewObjectsReport.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewObjectsReport.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3});
            this.listViewObjectsReport.FullRowSelect = true;
            this.listViewObjectsReport.GridLines = true;
            this.listViewObjectsReport.HideSelection = false;
            this.listViewObjectsReport.Location = new System.Drawing.Point(12, 12);
            this.listViewObjectsReport.Name = "listViewObjectsReport";
            this.listViewObjectsReport.Size = new System.Drawing.Size(640, 221);
            this.listViewObjectsReport.TabIndex = 17;
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
            // comboBoxAdmissions
            // 
            this.comboBoxAdmissions.FormattingEnabled = true;
            this.comboBoxAdmissions.Items.AddRange(new object[] {
            "±1%",
            "±2%",
            "±5%",
            "±10%",
            "±20%"});
            this.comboBoxAdmissions.Location = new System.Drawing.Point(18, 62);
            this.comboBoxAdmissions.Name = "comboBoxAdmissions";
            this.comboBoxAdmissions.Size = new System.Drawing.Size(162, 21);
            this.comboBoxAdmissions.TabIndex = 0;
            // 
            // buttonWriteObject
            // 
            this.buttonWriteObject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonWriteObject.Location = new System.Drawing.Point(196, 60);
            this.buttonWriteObject.Name = "buttonWriteObject";
            this.buttonWriteObject.Size = new System.Drawing.Size(421, 23);
            this.buttonWriteObject.TabIndex = 14;
            this.buttonWriteObject.Text = "Добавить в допуск";
            this.buttonWriteObject.UseVisualStyleBackColor = true;
            this.buttonWriteObject.Click += new System.EventHandler(this.buttonWriteObject_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.comboBoxAdmissions);
            this.groupBox1.Controls.Add(this.buttonWriteObject);
            this.groupBox1.Location = new System.Drawing.Point(12, 266);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(640, 100);
            this.groupBox1.TabIndex = 20;
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
            // buttonWriteToReport
            // 
            this.buttonWriteToReport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonWriteToReport.Location = new System.Drawing.Point(12, 372);
            this.buttonWriteToReport.Name = "buttonWriteToReport";
            this.buttonWriteToReport.Size = new System.Drawing.Size(640, 23);
            this.buttonWriteToReport.TabIndex = 18;
            this.buttonWriteToReport.Text = "Записать в отчет";
            this.buttonWriteToReport.UseVisualStyleBackColor = true;
            this.buttonWriteToReport.Click += new System.EventHandler(this.buttonWriteToReport_Click);
            // 
            // buttonDeleteAdmission
            // 
            this.buttonDeleteAdmission.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDeleteAdmission.Enabled = false;
            this.buttonDeleteAdmission.Location = new System.Drawing.Point(265, 241);
            this.buttonDeleteAdmission.Name = "buttonDeleteAdmission";
            this.buttonDeleteAdmission.Size = new System.Drawing.Size(387, 23);
            this.buttonDeleteAdmission.TabIndex = 21;
            this.buttonDeleteAdmission.Text = "Удалить допуск";
            this.buttonDeleteAdmission.UseVisualStyleBackColor = true;
            this.buttonDeleteAdmission.Click += new System.EventHandler(this.buttonDeleteAdmission_Click);
            // 
            // AddAdmissionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(663, 402);
            this.Controls.Add(this.buttonDeleteAdmission);
            this.Controls.Add(this.listViewObjectsReport);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonWriteToReport);
            this.MinimumSize = new System.Drawing.Size(497, 441);
            this.Name = "AddAdmissionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Добавление значений допусков";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.AddAdmissionForm_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        public System.Windows.Forms.ListView listViewObjectsReport;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        public System.Windows.Forms.ComboBox comboBoxAdmissions;
        private System.Windows.Forms.Button buttonWriteObject;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonWriteToReport;
        private System.Windows.Forms.Button buttonDeleteAdmission;
    }
}