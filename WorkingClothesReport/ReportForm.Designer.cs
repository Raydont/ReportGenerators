namespace WorkingClothesReport
{
    partial class ReportForm
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
            this.typeGroupBox = new System.Windows.Forms.GroupBox();
            this.departRadioButton = new System.Windows.Forms.RadioButton();
            this.overallRadioButton = new System.Windows.Forms.RadioButton();
            this.periodGroupBox = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.halfYearComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.yearComboBox = new System.Windows.Forms.ComboBox();
            this.departComboBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.createButton = new System.Windows.Forms.Button();
            this.typeGroupBox.SuspendLayout();
            this.periodGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // typeGroupBox
            // 
            this.typeGroupBox.Controls.Add(this.departRadioButton);
            this.typeGroupBox.Controls.Add(this.overallRadioButton);
            this.typeGroupBox.Location = new System.Drawing.Point(12, 12);
            this.typeGroupBox.Name = "typeGroupBox";
            this.typeGroupBox.Size = new System.Drawing.Size(325, 62);
            this.typeGroupBox.TabIndex = 0;
            this.typeGroupBox.TabStop = false;
            this.typeGroupBox.Text = "Тип отчета";
            // 
            // departRadioButton
            // 
            this.departRadioButton.AutoSize = true;
            this.departRadioButton.Location = new System.Drawing.Point(175, 25);
            this.departRadioButton.Name = "departRadioButton";
            this.departRadioButton.Size = new System.Drawing.Size(143, 17);
            this.departRadioButton.TabIndex = 1;
            this.departRadioButton.TabStop = true;
            this.departRadioButton.Text = "Заявка подразделения";
            this.departRadioButton.UseVisualStyleBackColor = true;
            this.departRadioButton.CheckedChanged += new System.EventHandler(this.departRadioButton_CheckedChanged);
            // 
            // overallRadioButton
            // 
            this.overallRadioButton.AutoSize = true;
            this.overallRadioButton.Location = new System.Drawing.Point(20, 25);
            this.overallRadioButton.Name = "overallRadioButton";
            this.overallRadioButton.Size = new System.Drawing.Size(126, 17);
            this.overallRadioButton.TabIndex = 0;
            this.overallRadioButton.TabStop = true;
            this.overallRadioButton.Text = "Сводная ведомость";
            this.overallRadioButton.UseVisualStyleBackColor = true;
            this.overallRadioButton.CheckedChanged += new System.EventHandler(this.overallRadioButton_CheckedChanged);
            // 
            // periodGroupBox
            // 
            this.periodGroupBox.Controls.Add(this.label2);
            this.periodGroupBox.Controls.Add(this.halfYearComboBox);
            this.periodGroupBox.Controls.Add(this.label1);
            this.periodGroupBox.Controls.Add(this.yearComboBox);
            this.periodGroupBox.Location = new System.Drawing.Point(12, 80);
            this.periodGroupBox.Name = "periodGroupBox";
            this.periodGroupBox.Size = new System.Drawing.Size(325, 62);
            this.periodGroupBox.TabIndex = 1;
            this.periodGroupBox.TabStop = false;
            this.periodGroupBox.Text = "За период";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(256, 29);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Полугодие";
            // 
            // halfYearComboBox
            // 
            this.halfYearComboBox.FormattingEnabled = true;
            this.halfYearComboBox.Items.AddRange(new object[] {
            "1",
            "2"});
            this.halfYearComboBox.Location = new System.Drawing.Point(160, 26);
            this.halfYearComboBox.Name = "halfYearComboBox";
            this.halfYearComboBox.Size = new System.Drawing.Size(90, 21);
            this.halfYearComboBox.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(116, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(25, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Год";
            // 
            // yearComboBox
            // 
            this.yearComboBox.FormattingEnabled = true;
            this.yearComboBox.Location = new System.Drawing.Point(20, 26);
            this.yearComboBox.Name = "yearComboBox";
            this.yearComboBox.Size = new System.Drawing.Size(90, 21);
            this.yearComboBox.TabIndex = 0;
            // 
            // departComboBox
            // 
            this.departComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.departComboBox.FormattingEnabled = true;
            this.departComboBox.Location = new System.Drawing.Point(118, 159);
            this.departComboBox.Name = "departComboBox";
            this.departComboBox.Size = new System.Drawing.Size(219, 21);
            this.departComboBox.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 162);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Подразделение";
            // 
            // createButton
            // 
            this.createButton.Location = new System.Drawing.Point(110, 203);
            this.createButton.Name = "createButton";
            this.createButton.Size = new System.Drawing.Size(129, 33);
            this.createButton.TabIndex = 4;
            this.createButton.Text = "Сформировать";
            this.createButton.UseVisualStyleBackColor = true;
            this.createButton.Click += new System.EventHandler(this.createButton_Click);
            // 
            // ReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(349, 248);
            this.Controls.Add(this.createButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.departComboBox);
            this.Controls.Add(this.periodGroupBox);
            this.Controls.Add(this.typeGroupBox);
            this.Name = "ReportForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет по заявкам на спецодежду";
            this.Load += new System.EventHandler(this.ReportForm_Load);
            this.typeGroupBox.ResumeLayout(false);
            this.typeGroupBox.PerformLayout();
            this.periodGroupBox.ResumeLayout(false);
            this.periodGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox typeGroupBox;
        private System.Windows.Forms.RadioButton departRadioButton;
        private System.Windows.Forms.RadioButton overallRadioButton;
        private System.Windows.Forms.GroupBox periodGroupBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox halfYearComboBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox yearComboBox;
        private System.Windows.Forms.ComboBox departComboBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button createButton;
    }
}