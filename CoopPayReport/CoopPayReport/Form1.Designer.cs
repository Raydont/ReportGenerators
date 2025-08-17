namespace CoopPayReport
{
    partial class mainPeriodForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.dtpYear = new System.Windows.Forms.DateTimePicker();
            this.yearGroupBox = new System.Windows.Forms.GroupBox();
            this.calendarGroupBox = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpEndPeriod = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpStartPeriod = new System.Windows.Forms.DateTimePicker();
            this.calendarCheckBox = new System.Windows.Forms.CheckBox();
            this.reportButton = new System.Windows.Forms.Button();
            this.yearGroupBox.SuspendLayout();
            this.calendarGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // dtpYear
            // 
            this.dtpYear.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpYear.Location = new System.Drawing.Point(90, 23);
            this.dtpYear.Name = "dtpYear";
            this.dtpYear.Size = new System.Drawing.Size(134, 20);
            this.dtpYear.TabIndex = 0;
            // 
            // yearGroupBox
            // 
            this.yearGroupBox.Controls.Add(this.dtpYear);
            this.yearGroupBox.Location = new System.Drawing.Point(12, 38);
            this.yearGroupBox.Name = "yearGroupBox";
            this.yearGroupBox.Size = new System.Drawing.Size(303, 60);
            this.yearGroupBox.TabIndex = 1;
            this.yearGroupBox.TabStop = false;
            this.yearGroupBox.Text = "Год";
            // 
            // calendarGroupBox
            // 
            this.calendarGroupBox.Controls.Add(this.label2);
            this.calendarGroupBox.Controls.Add(this.dtpEndPeriod);
            this.calendarGroupBox.Controls.Add(this.label1);
            this.calendarGroupBox.Controls.Add(this.dtpStartPeriod);
            this.calendarGroupBox.Enabled = false;
            this.calendarGroupBox.Location = new System.Drawing.Point(12, 104);
            this.calendarGroupBox.Name = "calendarGroupBox";
            this.calendarGroupBox.Size = new System.Drawing.Size(303, 91);
            this.calendarGroupBox.TabIndex = 2;
            this.calendarGroupBox.TabStop = false;
            this.calendarGroupBox.Text = "Календарный период";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(56, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(19, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "по";
            // 
            // dtpEndPeriod
            // 
            this.dtpEndPeriod.Location = new System.Drawing.Point(89, 60);
            this.dtpEndPeriod.Name = "dtpEndPeriod";
            this.dtpEndPeriod.Size = new System.Drawing.Size(134, 20);
            this.dtpEndPeriod.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(59, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(14, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "С";
            // 
            // dtpStartPeriod
            // 
            this.dtpStartPeriod.Location = new System.Drawing.Point(89, 24);
            this.dtpStartPeriod.Name = "dtpStartPeriod";
            this.dtpStartPeriod.Size = new System.Drawing.Size(134, 20);
            this.dtpStartPeriod.TabIndex = 0;
            // 
            // calendarCheckBox
            // 
            this.calendarCheckBox.AutoSize = true;
            this.calendarCheckBox.Location = new System.Drawing.Point(103, 16);
            this.calendarCheckBox.Name = "calendarCheckBox";
            this.calendarCheckBox.Size = new System.Drawing.Size(134, 17);
            this.calendarCheckBox.TabIndex = 3;
            this.calendarCheckBox.Text = "Календарный период";
            this.calendarCheckBox.UseVisualStyleBackColor = true;
            this.calendarCheckBox.CheckedChanged += new System.EventHandler(this.calendarCheckBox_CheckedChanged);
            // 
            // reportButton
            // 
            this.reportButton.Location = new System.Drawing.Point(113, 208);
            this.reportButton.Name = "reportButton";
            this.reportButton.Size = new System.Drawing.Size(105, 23);
            this.reportButton.TabIndex = 4;
            this.reportButton.Text = "Сформировать";
            this.reportButton.UseVisualStyleBackColor = true;
            this.reportButton.Click += new System.EventHandler(this.reportButton_Click);
            // 
            // mainPeriodForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(327, 244);
            this.Controls.Add(this.reportButton);
            this.Controls.Add(this.calendarCheckBox);
            this.Controls.Add(this.calendarGroupBox);
            this.Controls.Add(this.yearGroupBox);
            this.Name = "mainPeriodForm";
            this.Text = "Выберите период";
            this.yearGroupBox.ResumeLayout(false);
            this.calendarGroupBox.ResumeLayout(false);
            this.calendarGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtpYear;
        private System.Windows.Forms.GroupBox yearGroupBox;
        private System.Windows.Forms.GroupBox calendarGroupBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpEndPeriod;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpStartPeriod;
        private System.Windows.Forms.CheckBox calendarCheckBox;
        private System.Windows.Forms.Button reportButton;
    }
}

