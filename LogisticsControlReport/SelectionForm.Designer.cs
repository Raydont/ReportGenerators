namespace LogisticsControlReport
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
            this.radioButtonSelectedPeriodsFails = new System.Windows.Forms.RadioButton();
            this.radioButtonSelectedPeriod = new System.Windows.Forms.RadioButton();
            this.radioButtonAllTime = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTimePickerBeginFails = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerBegin = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEndFails = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.comboBoxMaker = new System.Windows.Forms.ComboBox();
            this.radioButtonSelectedMaker = new System.Windows.Forms.RadioButton();
            this.radioButtonAllMaker = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.comboBoxDateExecute = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(16, 428);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(509, 23);
            this.buttonMakeReport.TabIndex = 1;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click_1);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonSelectedPeriodsFails);
            this.groupBox1.Controls.Add(this.radioButtonSelectedPeriod);
            this.groupBox1.Controls.Add(this.radioButtonAllTime);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.dateTimePickerBeginFails);
            this.groupBox1.Controls.Add(this.dateTimePickerBegin);
            this.groupBox1.Controls.Add(this.dateTimePickerEndFails);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.dateTimePickerEnd);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Location = new System.Drawing.Point(16, 18);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(509, 183);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Выбор периода";
            // 
            // radioButtonSelectedPeriodsFails
            // 
            this.radioButtonSelectedPeriodsFails.AutoSize = true;
            this.radioButtonSelectedPeriodsFails.Location = new System.Drawing.Point(22, 125);
            this.radioButtonSelectedPeriodsFails.Name = "radioButtonSelectedPeriodsFails";
            this.radioButtonSelectedPeriodsFails.Size = new System.Drawing.Size(263, 17);
            this.radioButtonSelectedPeriodsFails.TabIndex = 19;
            this.radioButtonSelectedPeriodsFails.Text = "За выбранный период выставления незачетов";
            this.radioButtonSelectedPeriodsFails.UseVisualStyleBackColor = true;
            this.radioButtonSelectedPeriodsFails.CheckedChanged += new System.EventHandler(this.radioButtonSelectedPeriodsFails_CheckedChanged);
            // 
            // radioButtonSelectedPeriod
            // 
            this.radioButtonSelectedPeriod.AutoSize = true;
            this.radioButtonSelectedPeriod.Checked = true;
            this.radioButtonSelectedPeriod.Location = new System.Drawing.Point(23, 60);
            this.radioButtonSelectedPeriod.Name = "radioButtonSelectedPeriod";
            this.radioButtonSelectedPeriod.Size = new System.Drawing.Size(201, 17);
            this.radioButtonSelectedPeriod.TabIndex = 19;
            this.radioButtonSelectedPeriod.TabStop = true;
            this.radioButtonSelectedPeriod.Text = "За выбранный период исполнения";
            this.radioButtonSelectedPeriod.UseVisualStyleBackColor = true;
            this.radioButtonSelectedPeriod.CheckedChanged += new System.EventHandler(this.radioButtonSelectedPeriod_CheckedChanged);
            // 
            // radioButtonAllTime
            // 
            this.radioButtonAllTime.AutoSize = true;
            this.radioButtonAllTime.Location = new System.Drawing.Point(23, 29);
            this.radioButtonAllTime.Name = "radioButtonAllTime";
            this.radioButtonAllTime.Size = new System.Drawing.Size(94, 17);
            this.radioButtonAllTime.TabIndex = 18;
            this.radioButtonAllTime.Text = "За все время";
            this.radioButtonAllTime.UseVisualStyleBackColor = true;
            this.radioButtonAllTime.CheckedChanged += new System.EventHandler(this.radioButtonAllTime_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(301, 152);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(19, 13);
            this.label1.TabIndex = 17;
            this.label1.Text = "по";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(305, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(19, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "по";
            // 
            // dateTimePickerBeginFails
            // 
            this.dateTimePickerBeginFails.Enabled = false;
            this.dateTimePickerBeginFails.Location = new System.Drawing.Point(333, 121);
            this.dateTimePickerBeginFails.Name = "dateTimePickerBeginFails";
            this.dateTimePickerBeginFails.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerBeginFails.TabIndex = 14;
            // 
            // dateTimePickerBegin
            // 
            this.dateTimePickerBegin.Location = new System.Drawing.Point(333, 51);
            this.dateTimePickerBegin.Name = "dateTimePickerBegin";
            this.dateTimePickerBegin.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerBegin.TabIndex = 14;
            // 
            // dateTimePickerEndFails
            // 
            this.dateTimePickerEndFails.Enabled = false;
            this.dateTimePickerEndFails.Location = new System.Drawing.Point(333, 148);
            this.dateTimePickerEndFails.Name = "dateTimePickerEndFails";
            this.dateTimePickerEndFails.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerEndFails.TabIndex = 15;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(302, 124);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(13, 13);
            this.label2.TabIndex = 16;
            this.label2.Text = "c";
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(333, 78);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerEnd.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(306, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(13, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "c";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.comboBoxMaker);
            this.groupBox2.Controls.Add(this.radioButtonSelectedMaker);
            this.groupBox2.Controls.Add(this.radioButtonAllMaker);
            this.groupBox2.Location = new System.Drawing.Point(16, 221);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(509, 86);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Выбор исполнителя";
            // 
            // comboBoxMaker
            // 
            this.comboBoxMaker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxMaker.Enabled = false;
            this.comboBoxMaker.FormattingEnabled = true;
            this.comboBoxMaker.Location = new System.Drawing.Point(164, 48);
            this.comboBoxMaker.Name = "comboBoxMaker";
            this.comboBoxMaker.Size = new System.Drawing.Size(330, 21);
            this.comboBoxMaker.TabIndex = 8;
            // 
            // radioButtonSelectedMaker
            // 
            this.radioButtonSelectedMaker.AutoSize = true;
            this.radioButtonSelectedMaker.Location = new System.Drawing.Point(164, 25);
            this.radioButtonSelectedMaker.Name = "radioButtonSelectedMaker";
            this.radioButtonSelectedMaker.Size = new System.Drawing.Size(152, 17);
            this.radioButtonSelectedMaker.TabIndex = 19;
            this.radioButtonSelectedMaker.Text = "Выбранный исполнитель";
            this.radioButtonSelectedMaker.UseVisualStyleBackColor = true;
            this.radioButtonSelectedMaker.CheckedChanged += new System.EventHandler(this.radioButtonSelectedMaker_CheckedChanged);
            // 
            // radioButtonAllMaker
            // 
            this.radioButtonAllMaker.AutoSize = true;
            this.radioButtonAllMaker.Checked = true;
            this.radioButtonAllMaker.Location = new System.Drawing.Point(22, 25);
            this.radioButtonAllMaker.Name = "radioButtonAllMaker";
            this.radioButtonAllMaker.Size = new System.Drawing.Size(112, 17);
            this.radioButtonAllMaker.TabIndex = 18;
            this.radioButtonAllMaker.TabStop = true;
            this.radioButtonAllMaker.Text = "Все исполнители";
            this.radioButtonAllMaker.UseVisualStyleBackColor = true;
            this.radioButtonAllMaker.CheckedChanged += new System.EventHandler(this.radioButtonAllMaker_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.comboBoxDateExecute);
            this.groupBox3.Location = new System.Drawing.Point(16, 326);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(509, 81);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Выбор даты исполнения";
            // 
            // comboBoxDateExecute
            // 
            this.comboBoxDateExecute.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxDateExecute.FormattingEnabled = true;
            this.comboBoxDateExecute.Items.AddRange(new object[] {
            "все",
            "только просроченные",
            "только в работе/на доработке",
            "только выполненные"});
            this.comboBoxDateExecute.Location = new System.Drawing.Point(84, 33);
            this.comboBoxDateExecute.Name = "comboBoxDateExecute";
            this.comboBoxDateExecute.Size = new System.Drawing.Size(330, 21);
            this.comboBoxDateExecute.TabIndex = 8;
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(548, 475);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonMakeReport);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройка отчета \"Контроль обработки и исполнения документов ОМТС\"";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dateTimePickerBegin;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RadioButton radioButtonSelectedPeriod;
        private System.Windows.Forms.RadioButton radioButtonAllTime;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox comboBoxMaker;
        private System.Windows.Forms.RadioButton radioButtonSelectedMaker;
        private System.Windows.Forms.RadioButton radioButtonAllMaker;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ComboBox comboBoxDateExecute;
        private System.Windows.Forms.RadioButton radioButtonSelectedPeriodsFails;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePickerBeginFails;
        private System.Windows.Forms.DateTimePicker dateTimePickerEndFails;
        private System.Windows.Forms.Label label2;
    }
}
