namespace CompletionDatesContractsReport
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
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comboBoxSpecificMaker = new System.Windows.Forms.ComboBox();
            this.comboBoxSpecificDepartament = new System.Windows.Forms.ComboBox();
            this.radioButtonSpecificMaker = new System.Windows.Forms.RadioButton();
            this.radioButtonSpecificDepartament = new System.Windows.Forms.RadioButton();
            this.radioButtonAllJournal = new System.Windows.Forms.RadioButton();
            this.comboBoxOrderSum = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(9, 168);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(375, 23);
            this.buttonMakeReport.TabIndex = 4;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.comboBoxOrderSum);
            this.groupBox1.Controls.Add(this.comboBoxSpecificMaker);
            this.groupBox1.Controls.Add(this.comboBoxSpecificDepartament);
            this.groupBox1.Controls.Add(this.radioButtonSpecificMaker);
            this.groupBox1.Controls.Add(this.radioButtonSpecificDepartament);
            this.groupBox1.Controls.Add(this.radioButtonAllJournal);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(370, 147);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Отчет для";
            // 
            // comboBoxSpecificMaker
            // 
            this.comboBoxSpecificMaker.Enabled = false;
            this.comboBoxSpecificMaker.FormattingEnabled = true;
            this.comboBoxSpecificMaker.Location = new System.Drawing.Point(128, 85);
            this.comboBoxSpecificMaker.Name = "comboBoxSpecificMaker";
            this.comboBoxSpecificMaker.Size = new System.Drawing.Size(236, 21);
            this.comboBoxSpecificMaker.TabIndex = 4;
            // 
            // comboBoxSpecificDepartament
            // 
            this.comboBoxSpecificDepartament.Enabled = false;
            this.comboBoxSpecificDepartament.FormattingEnabled = true;
            this.comboBoxSpecificDepartament.Location = new System.Drawing.Point(128, 53);
            this.comboBoxSpecificDepartament.Name = "comboBoxSpecificDepartament";
            this.comboBoxSpecificDepartament.Size = new System.Drawing.Size(236, 21);
            this.comboBoxSpecificDepartament.TabIndex = 3;
            // 
            // radioButtonSpecificMaker
            // 
            this.radioButtonSpecificMaker.AutoSize = true;
            this.radioButtonSpecificMaker.Location = new System.Drawing.Point(17, 86);
            this.radioButtonSpecificMaker.Name = "radioButtonSpecificMaker";
            this.radioButtonSpecificMaker.Size = new System.Drawing.Size(92, 17);
            this.radioButtonSpecificMaker.TabIndex = 2;
            this.radioButtonSpecificMaker.Text = "Исполнителя";
            this.radioButtonSpecificMaker.UseVisualStyleBackColor = true;
            this.radioButtonSpecificMaker.CheckedChanged += new System.EventHandler(this.radioButtonSpecificMaker_CheckedChanged);
            // 
            // radioButtonSpecificDepartament
            // 
            this.radioButtonSpecificDepartament.AutoSize = true;
            this.radioButtonSpecificDepartament.Location = new System.Drawing.Point(17, 54);
            this.radioButtonSpecificDepartament.Name = "radioButtonSpecificDepartament";
            this.radioButtonSpecificDepartament.Size = new System.Drawing.Size(105, 17);
            this.radioButtonSpecificDepartament.TabIndex = 1;
            this.radioButtonSpecificDepartament.Text = "Подразделения";
            this.radioButtonSpecificDepartament.UseVisualStyleBackColor = true;
            this.radioButtonSpecificDepartament.CheckedChanged += new System.EventHandler(this.radioButtonSpecificDepartament_CheckedChanged);
            // 
            // radioButtonAllJournal
            // 
            this.radioButtonAllJournal.AutoSize = true;
            this.radioButtonAllJournal.Checked = true;
            this.radioButtonAllJournal.Location = new System.Drawing.Point(17, 23);
            this.radioButtonAllJournal.Name = "radioButtonAllJournal";
            this.radioButtonAllJournal.Size = new System.Drawing.Size(128, 17);
            this.radioButtonAllJournal.TabIndex = 0;
            this.radioButtonAllJournal.TabStop = true;
            this.radioButtonAllJournal.Text = "Всего журнала ДЗУ";
            this.radioButtonAllJournal.UseVisualStyleBackColor = true;
            this.radioButtonAllJournal.CheckedChanged += new System.EventHandler(this.radioButtonAllJournal_CheckedChanged);
            // 
            // comboBoxOrderSum
            // 
            this.comboBoxOrderSum.FormattingEnabled = true;
            this.comboBoxOrderSum.Items.AddRange(new object[] {
            "от 100 тыс.руб.",
            "от 500 тыс.руб."});
            this.comboBoxOrderSum.Location = new System.Drawing.Point(128, 116);
            this.comboBoxOrderSum.Name = "comboBoxOrderSum";
            this.comboBoxOrderSum.Size = new System.Drawing.Size(236, 21);
            this.comboBoxOrderSum.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 119);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Сумма договора";
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(394, 201);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonMakeReport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет по срокам окончания действия ДЗУ";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox comboBoxSpecificMaker;
        private System.Windows.Forms.ComboBox comboBoxSpecificDepartament;
        private System.Windows.Forms.RadioButton radioButtonSpecificMaker;
        private System.Windows.Forms.RadioButton radioButtonSpecificDepartament;
        private System.Windows.Forms.RadioButton radioButtonAllJournal;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboBoxOrderSum;
    }
}