namespace CooperationReport
{
    partial class SelectionForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
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
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.makeReportButton = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerBegin = new System.Windows.Forms.DateTimePicker();
            this.radioButtonForPeriod = new System.Windows.Forms.RadioButton();
            this.comboBoxYear = new System.Windows.Forms.ComboBox();
            this.dateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.comboBoxMonth = new System.Windows.Forms.ComboBox();
            this.radioButtonDate = new System.Windows.Forms.RadioButton();
            this.radioButtonMonth = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.radButExpense = new System.Windows.Forms.RadioButton();
            this.radButArrival = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.radButContract = new System.Windows.Forms.RadioButton();
            this.radButDevice = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // makeReportButton
            // 
            this.makeReportButton.Location = new System.Drawing.Point(12, 191);
            this.makeReportButton.Name = "makeReportButton";
            this.makeReportButton.Size = new System.Drawing.Size(492, 23);
            this.makeReportButton.TabIndex = 7;
            this.makeReportButton.Text = "Создать отчет";
            this.makeReportButton.UseVisualStyleBackColor = true;
            this.makeReportButton.Click += new System.EventHandler(this.makeReportButton_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Файлы IPC|*.ipc";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.dateTimePickerEnd);
            this.groupBox1.Controls.Add(this.dateTimePickerBegin);
            this.groupBox1.Controls.Add(this.radioButtonForPeriod);
            this.groupBox1.Controls.Add(this.comboBoxYear);
            this.groupBox1.Controls.Add(this.dateTimePicker);
            this.groupBox1.Controls.Add(this.comboBoxMonth);
            this.groupBox1.Controls.Add(this.radioButtonDate);
            this.groupBox1.Controls.Add(this.radioButtonMonth);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(492, 102);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Вид отчета";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(279, 75);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(19, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "по";
            this.label2.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(279, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(13, 13);
            this.label1.TabIndex = 14;
            this.label1.Text = "с";
            this.label1.Visible = false;
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(302, 73);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(167, 20);
            this.dateTimePickerEnd.TabIndex = 13;
            this.dateTimePickerEnd.Visible = false;
            // 
            // dateTimePickerBegin
            // 
            this.dateTimePickerBegin.Location = new System.Drawing.Point(303, 43);
            this.dateTimePickerBegin.Name = "dateTimePickerBegin";
            this.dateTimePickerBegin.Size = new System.Drawing.Size(167, 20);
            this.dateTimePickerBegin.TabIndex = 12;
            this.dateTimePickerBegin.Visible = false;
            // 
            // radioButtonForPeriod
            // 
            this.radioButtonForPeriod.AutoSize = true;
            this.radioButtonForPeriod.Location = new System.Drawing.Point(334, 19);
            this.radioButtonForPeriod.Name = "radioButtonForPeriod";
            this.radioButtonForPeriod.Size = new System.Drawing.Size(77, 17);
            this.radioButtonForPeriod.TabIndex = 11;
            this.radioButtonForPeriod.Text = "За период";
            this.radioButtonForPeriod.UseVisualStyleBackColor = true;
            this.radioButtonForPeriod.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // comboBoxYear
            // 
            this.comboBoxYear.FormattingEnabled = true;
            this.comboBoxYear.Location = new System.Drawing.Point(169, 64);
            this.comboBoxYear.Name = "comboBoxYear";
            this.comboBoxYear.Size = new System.Drawing.Size(74, 21);
            this.comboBoxYear.TabIndex = 10;
            // 
            // dateTimePicker
            // 
            this.dateTimePicker.Location = new System.Drawing.Point(16, 64);
            this.dateTimePicker.Name = "dateTimePicker";
            this.dateTimePicker.Size = new System.Drawing.Size(227, 20);
            this.dateTimePicker.TabIndex = 9;
            this.dateTimePicker.Visible = false;
            // 
            // comboBoxMonth
            // 
            this.comboBoxMonth.FormattingEnabled = true;
            this.comboBoxMonth.Items.AddRange(new object[] {
            "Январь",
            "Февраль",
            "Март",
            "Апрель",
            "Май",
            "Июнь",
            "Июль",
            "Август",
            "Сентябрь",
            "Октябрь",
            "Ноябрь",
            "Декабрь"});
            this.comboBoxMonth.Location = new System.Drawing.Point(16, 64);
            this.comboBoxMonth.Name = "comboBoxMonth";
            this.comboBoxMonth.Size = new System.Drawing.Size(119, 21);
            this.comboBoxMonth.TabIndex = 9;
            // 
            // radioButtonDate
            // 
            this.radioButtonDate.AutoSize = true;
            this.radioButtonDate.Location = new System.Drawing.Point(169, 19);
            this.radioButtonDate.Name = "radioButtonDate";
            this.radioButtonDate.Size = new System.Drawing.Size(95, 17);
            this.radioButtonDate.TabIndex = 1;
            this.radioButtonDate.Text = "Отчет по дате";
            this.radioButtonDate.UseVisualStyleBackColor = true;
            this.radioButtonDate.CheckedChanged += new System.EventHandler(this.radioButtonDate_CheckedChanged);
            // 
            // radioButtonMonth
            // 
            this.radioButtonMonth.AutoSize = true;
            this.radioButtonMonth.Checked = true;
            this.radioButtonMonth.Location = new System.Drawing.Point(16, 20);
            this.radioButtonMonth.Name = "radioButtonMonth";
            this.radioButtonMonth.Size = new System.Drawing.Size(104, 17);
            this.radioButtonMonth.TabIndex = 0;
            this.radioButtonMonth.TabStop = true;
            this.radioButtonMonth.Text = "Отчет на месяц";
            this.radioButtonMonth.UseVisualStyleBackColor = true;
            this.radioButtonMonth.CheckedChanged += new System.EventHandler(this.radioButtonMonth_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.radButExpense);
            this.groupBox2.Controls.Add(this.radButArrival);
            this.groupBox2.Location = new System.Drawing.Point(12, 120);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(243, 54);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Тип данных";
            // 
            // radButExpense
            // 
            this.radButExpense.AutoSize = true;
            this.radButExpense.Checked = true;
            this.radButExpense.Location = new System.Drawing.Point(148, 23);
            this.radButExpense.Name = "radButExpense";
            this.radButExpense.Size = new System.Drawing.Size(61, 17);
            this.radButExpense.TabIndex = 10;
            this.radButExpense.TabStop = true;
            this.radButExpense.Text = "Расход";
            this.radButExpense.UseVisualStyleBackColor = true;
            // 
            // radButArrival
            // 
            this.radButArrival.AutoSize = true;
            this.radButArrival.Location = new System.Drawing.Point(16, 23);
            this.radButArrival.Name = "radButArrival";
            this.radButArrival.Size = new System.Drawing.Size(62, 17);
            this.radButArrival.TabIndex = 0;
            this.radButArrival.TabStop = true;
            this.radButArrival.Text = "Приход";
            this.radButArrival.UseVisualStyleBackColor = true;
            this.radButArrival.CheckedChanged += new System.EventHandler(this.radButArrival_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.radButContract);
            this.groupBox3.Controls.Add(this.radButDevice);
            this.groupBox3.Location = new System.Drawing.Point(261, 120);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(243, 54);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Суммирование объектов";
            // 
            // radButContract
            // 
            this.radButContract.AutoSize = true;
            this.radButContract.Location = new System.Drawing.Point(148, 23);
            this.radButContract.Name = "radButContract";
            this.radButContract.Size = new System.Drawing.Size(75, 17);
            this.radButContract.TabIndex = 10;
            this.radButContract.Text = "Договора";
            this.radButContract.UseVisualStyleBackColor = true;
            this.radButContract.CheckedChanged += new System.EventHandler(this.radButContract_CheckedChanged);
            // 
            // radButDevice
            // 
            this.radButDevice.AutoSize = true;
            this.radButDevice.Checked = true;
            this.radButDevice.Location = new System.Drawing.Point(16, 23);
            this.radButDevice.Name = "radButDevice";
            this.radButDevice.Size = new System.Drawing.Size(85, 17);
            this.radButDevice.TabIndex = 0;
            this.radButDevice.TabStop = true;
            this.radButDevice.Text = "Устройства";
            this.radButDevice.UseVisualStyleBackColor = true;
            this.radButDevice.CheckedChanged += new System.EventHandler(this.radButDevice_CheckedChanged);
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 237);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.makeReportButton);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет кооперации";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button makeReportButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker dateTimePicker;
        private System.Windows.Forms.ComboBox comboBoxMonth;
        private System.Windows.Forms.RadioButton radioButtonDate;
        private System.Windows.Forms.RadioButton radioButtonMonth;
        private System.Windows.Forms.ComboBox comboBoxYear;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton radButExpense;
        private System.Windows.Forms.RadioButton radButArrival;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton radButContract;
        private System.Windows.Forms.RadioButton radButDevice;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.DateTimePicker dateTimePickerBegin;
        private System.Windows.Forms.RadioButton radioButtonForPeriod;
    }
}

