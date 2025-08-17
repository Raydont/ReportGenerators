namespace RunningContractsReport
{
    partial class SelectionForm
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
            this.radButDOD = new System.Windows.Forms.RadioButton();
            this.radButDZU = new System.Windows.Forms.RadioButton();
            this.checkBoxActualContract = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerBegin = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.buttonCansel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // radButDOD
            // 
            this.radButDOD.AutoSize = true;
            this.radButDOD.Checked = true;
            this.radButDOD.Location = new System.Drawing.Point(6, 19);
            this.radButDOD.Name = "radButDOD";
            this.radButDOD.Size = new System.Drawing.Size(201, 17);
            this.radButDOD.TabIndex = 0;
            this.radButDOD.TabStop = true;
            this.radButDOD.Text = "Договоры основной деятельности";
            this.radButDOD.UseVisualStyleBackColor = true;
            this.radButDOD.CheckedChanged += new System.EventHandler(this.radButDOD_CheckedChanged);
            // 
            // radButDZU
            // 
            this.radButDZU.AutoSize = true;
            this.radButDZU.Location = new System.Drawing.Point(313, 19);
            this.radButDZU.Name = "radButDZU";
            this.radButDZU.Size = new System.Drawing.Size(166, 17);
            this.radButDZU.TabIndex = 1;
            this.radButDZU.Text = "Договоры закупки и услуги";
            this.radButDZU.UseVisualStyleBackColor = true;
            this.radButDZU.CheckedChanged += new System.EventHandler(this.radButDZU_CheckedChanged);
            // 
            // checkBoxActualContract
            // 
            this.checkBoxActualContract.AutoSize = true;
            this.checkBoxActualContract.Checked = true;
            this.checkBoxActualContract.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxActualContract.Location = new System.Drawing.Point(32, 82);
            this.checkBoxActualContract.Name = "checkBoxActualContract";
            this.checkBoxActualContract.Size = new System.Drawing.Size(150, 17);
            this.checkBoxActualContract.TabIndex = 3;
            this.checkBoxActualContract.Text = "Действующие договоры";
            this.checkBoxActualContract.UseVisualStyleBackColor = true;
            this.checkBoxActualContract.CheckedChanged += new System.EventHandler(this.checkBoxActualContract_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radButDOD);
            this.groupBox1.Controls.Add(this.radButDZU);
            this.groupBox1.Location = new System.Drawing.Point(26, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(515, 57);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Вид договоров";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dateTimePickerEnd);
            this.groupBox2.Controls.Add(this.dateTimePickerBegin);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Enabled = false;
            this.groupBox2.Location = new System.Drawing.Point(26, 105);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(515, 97);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Ввод диапозона дат отчета ";
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(172, 58);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(156, 20);
            this.dateTimePickerEnd.TabIndex = 9;
            // 
            // dateTimePickerBegin
            // 
            this.dateTimePickerBegin.Location = new System.Drawing.Point(172, 29);
            this.dateTimePickerBegin.Name = "dateTimePickerBegin";
            this.dateTimePickerBegin.Size = new System.Drawing.Size(156, 20);
            this.dateTimePickerBegin.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(142, 64);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(21, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "По";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(142, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(14, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "С";
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(426, 211);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(121, 23);
            this.buttonMakeReport.TabIndex = 8;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // buttonCansel
            // 
            this.buttonCansel.Location = new System.Drawing.Point(345, 211);
            this.buttonCansel.Name = "buttonCansel";
            this.buttonCansel.Size = new System.Drawing.Size(75, 23);
            this.buttonCansel.TabIndex = 9;
            this.buttonCansel.Text = "Отмена";
            this.buttonCansel.UseVisualStyleBackColor = true;
            this.buttonCansel.Click += new System.EventHandler(this.buttonCansel_Click);
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 243);
            this.Controls.Add(this.buttonCansel);
            this.Controls.Add(this.buttonMakeReport);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.checkBoxActualContract);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Параметры отчета";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.RadioButton radButDOD;
        public System.Windows.Forms.RadioButton radButDZU;
        public System.Windows.Forms.CheckBox checkBoxActualContract;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.Button buttonCansel;
        private System.Windows.Forms.DateTimePicker dateTimePickerBegin;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
    }
}