namespace OfficeEquipmentReport
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
            this.comboBoxDepartments = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBoxComputer = new System.Windows.Forms.CheckBox();
            this.checkBoxMonitor = new System.Windows.Forms.CheckBox();
            this.checkBoxPrinter = new System.Windows.Forms.CheckBox();
            this.checkBoxScanner = new System.Windows.Forms.CheckBox();
            this.checkBoxNotebook = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBoxMFU = new System.Windows.Forms.CheckBox();
            this.checkBoxOtherDevices = new System.Windows.Forms.CheckBox();
            this.checkBoxStorage = new System.Windows.Forms.CheckBox();
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboBoxDepartments
            // 
            this.comboBoxDepartments.FormattingEnabled = true;
            this.comboBoxDepartments.Location = new System.Drawing.Point(114, 25);
            this.comboBoxDepartments.Name = "comboBoxDepartments";
            this.comboBoxDepartments.Size = new System.Drawing.Size(184, 21);
            this.comboBoxDepartments.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Подразделение";
            // 
            // checkBoxComputer
            // 
            this.checkBoxComputer.AutoSize = true;
            this.checkBoxComputer.Checked = true;
            this.checkBoxComputer.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxComputer.Location = new System.Drawing.Point(24, 78);
            this.checkBoxComputer.Name = "checkBoxComputer";
            this.checkBoxComputer.Size = new System.Drawing.Size(111, 17);
            this.checkBoxComputer.TabIndex = 2;
            this.checkBoxComputer.Text = "Системный блок";
            this.checkBoxComputer.UseVisualStyleBackColor = true;
            // 
            // checkBoxMonitor
            // 
            this.checkBoxMonitor.AutoSize = true;
            this.checkBoxMonitor.Checked = true;
            this.checkBoxMonitor.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxMonitor.Location = new System.Drawing.Point(24, 103);
            this.checkBoxMonitor.Name = "checkBoxMonitor";
            this.checkBoxMonitor.Size = new System.Drawing.Size(70, 17);
            this.checkBoxMonitor.TabIndex = 3;
            this.checkBoxMonitor.Text = "Монитор";
            this.checkBoxMonitor.UseVisualStyleBackColor = true;
            // 
            // checkBoxPrinter
            // 
            this.checkBoxPrinter.AutoSize = true;
            this.checkBoxPrinter.Checked = true;
            this.checkBoxPrinter.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxPrinter.Location = new System.Drawing.Point(185, 103);
            this.checkBoxPrinter.Name = "checkBoxPrinter";
            this.checkBoxPrinter.Size = new System.Drawing.Size(69, 17);
            this.checkBoxPrinter.TabIndex = 4;
            this.checkBoxPrinter.Text = "Принтер";
            this.checkBoxPrinter.UseVisualStyleBackColor = true;
            // 
            // checkBoxScanner
            // 
            this.checkBoxScanner.AutoSize = true;
            this.checkBoxScanner.Checked = true;
            this.checkBoxScanner.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxScanner.Location = new System.Drawing.Point(185, 128);
            this.checkBoxScanner.Name = "checkBoxScanner";
            this.checkBoxScanner.Size = new System.Drawing.Size(63, 17);
            this.checkBoxScanner.TabIndex = 5;
            this.checkBoxScanner.Text = "Сканер";
            this.checkBoxScanner.UseVisualStyleBackColor = true;
            // 
            // checkBoxNotebook
            // 
            this.checkBoxNotebook.AutoSize = true;
            this.checkBoxNotebook.Checked = true;
            this.checkBoxNotebook.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxNotebook.Location = new System.Drawing.Point(24, 128);
            this.checkBoxNotebook.Name = "checkBoxNotebook";
            this.checkBoxNotebook.Size = new System.Drawing.Size(67, 17);
            this.checkBoxNotebook.TabIndex = 8;
            this.checkBoxNotebook.Text = "Ноутбук";
            this.checkBoxNotebook.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBoxMFU);
            this.groupBox1.Controls.Add(this.comboBoxDepartments);
            this.groupBox1.Controls.Add(this.checkBoxOtherDevices);
            this.groupBox1.Controls.Add(this.checkBoxStorage);
            this.groupBox1.Controls.Add(this.checkBoxNotebook);
            this.groupBox1.Controls.Add(this.checkBoxScanner);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.checkBoxComputer);
            this.groupBox1.Controls.Add(this.checkBoxMonitor);
            this.groupBox1.Controls.Add(this.checkBoxPrinter);
            this.groupBox1.Location = new System.Drawing.Point(19, 16);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(321, 160);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Настройки отчета";
            // 
            // checkBoxMFU
            // 
            this.checkBoxMFU.AutoSize = true;
            this.checkBoxMFU.Checked = true;
            this.checkBoxMFU.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxMFU.Location = new System.Drawing.Point(185, 78);
            this.checkBoxMFU.Name = "checkBoxMFU";
            this.checkBoxMFU.Size = new System.Drawing.Size(54, 17);
            this.checkBoxMFU.TabIndex = 9;
            this.checkBoxMFU.Text = "МФУ";
            this.checkBoxMFU.UseVisualStyleBackColor = true;
            // 
            // checkBoxOtherDevices
            // 
            this.checkBoxOtherDevices.AutoSize = true;
            this.checkBoxOtherDevices.Location = new System.Drawing.Point(275, 128);
            this.checkBoxOtherDevices.Name = "checkBoxOtherDevices";
            this.checkBoxOtherDevices.Size = new System.Drawing.Size(123, 17);
            this.checkBoxOtherDevices.TabIndex = 7;
            this.checkBoxOtherDevices.Text = "Прочие устройства";
            this.checkBoxOtherDevices.UseVisualStyleBackColor = true;
            this.checkBoxOtherDevices.Visible = false;
            // 
            // checkBoxStorage
            // 
            this.checkBoxStorage.AutoSize = true;
            this.checkBoxStorage.Location = new System.Drawing.Point(275, 103);
            this.checkBoxStorage.Name = "checkBoxStorage";
            this.checkBoxStorage.Size = new System.Drawing.Size(87, 17);
            this.checkBoxStorage.TabIndex = 6;
            this.checkBoxStorage.Text = "Накопитель";
            this.checkBoxStorage.UseVisualStyleBackColor = true;
            this.checkBoxStorage.Visible = false;
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(19, 192);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(321, 23);
            this.buttonMakeReport.TabIndex = 10;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(358, 229);
            this.Controls.Add(this.buttonMakeReport);
            this.Controls.Add(this.groupBox1);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Учет оргтехники по подразделениям";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxDepartments;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBoxComputer;
        private System.Windows.Forms.CheckBox checkBoxMonitor;
        private System.Windows.Forms.CheckBox checkBoxPrinter;
        private System.Windows.Forms.CheckBox checkBoxScanner;
        private System.Windows.Forms.CheckBox checkBoxNotebook;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.CheckBox checkBoxMFU;
        private System.Windows.Forms.CheckBox checkBoxOtherDevices;
        private System.Windows.Forms.CheckBox checkBoxStorage;
    }
}