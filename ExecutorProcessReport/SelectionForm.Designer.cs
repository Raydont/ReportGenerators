namespace ExecutorProcessReport
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
            this.comboBoxProcess = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBoxState = new System.Windows.Forms.ComboBox();
            this.groupBoxSelectPeriod = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTimePickerBegin = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBoxSelectObjects = new System.Windows.Forms.GroupBox();
            this.comboBoxAuthor = new System.Windows.Forms.ComboBox();
            this.comboBoxDepartment = new System.Windows.Forms.ComboBox();
            this.radioButtonObjectsByAuthor = new System.Windows.Forms.RadioButton();
            this.radioButtonObjectsByDepartment = new System.Windows.Forms.RadioButton();
            this.radioButtonAllObjects = new System.Windows.Forms.RadioButton();
            this.checkBoxIsState = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.NumericUDCountControlDays = new System.Windows.Forms.NumericUpDown();
            this.radioButtonChiefs = new System.Windows.Forms.RadioButton();
            this.groupBoxSelectPeriod.SuspendLayout();
            this.groupBoxSelectObjects.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NumericUDCountControlDays)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(14, 283);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(492, 23);
            this.buttonMakeReport.TabIndex = 1;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click_1);
            // 
            // comboBoxProcess
            // 
            this.comboBoxProcess.FormattingEnabled = true;
            this.comboBoxProcess.Location = new System.Drawing.Point(130, 14);
            this.comboBoxProcess.Name = "comboBoxProcess";
            this.comboBoxProcess.Size = new System.Drawing.Size(369, 21);
            this.comboBoxProcess.TabIndex = 2;
            this.comboBoxProcess.SelectedIndexChanged += new System.EventHandler(this.comboBoxProcess_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Процесс";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(112, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Состояния процесса";
            // 
            // comboBoxState
            // 
            this.comboBoxState.FormattingEnabled = true;
            this.comboBoxState.Location = new System.Drawing.Point(130, 43);
            this.comboBoxState.Name = "comboBoxState";
            this.comboBoxState.Size = new System.Drawing.Size(369, 21);
            this.comboBoxState.TabIndex = 4;
            this.comboBoxState.SelectedIndexChanged += new System.EventHandler(this.comboBoxState_SelectedIndexChanged);
            this.comboBoxState.SelectedValueChanged += new System.EventHandler(this.comboBoxState_SelectedIndexChanged);
            // 
            // groupBoxSelectPeriod
            // 
            this.groupBoxSelectPeriod.Controls.Add(this.label3);
            this.groupBoxSelectPeriod.Controls.Add(this.dateTimePickerBegin);
            this.groupBoxSelectPeriod.Controls.Add(this.dateTimePickerEnd);
            this.groupBoxSelectPeriod.Controls.Add(this.label4);
            this.groupBoxSelectPeriod.Location = new System.Drawing.Point(13, 186);
            this.groupBoxSelectPeriod.Name = "groupBoxSelectPeriod";
            this.groupBoxSelectPeriod.Size = new System.Drawing.Size(489, 71);
            this.groupBoxSelectPeriod.TabIndex = 6;
            this.groupBoxSelectPeriod.TabStop = false;
            this.groupBoxSelectPeriod.Text = "Выбор периода";
            this.groupBoxSelectPeriod.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(245, 33);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(19, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "по";
            // 
            // dateTimePickerBegin
            // 
            this.dateTimePickerBegin.Location = new System.Drawing.Point(90, 29);
            this.dateTimePickerBegin.Name = "dateTimePickerBegin";
            this.dateTimePickerBegin.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerBegin.TabIndex = 14;
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(277, 29);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerEnd.TabIndex = 15;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(59, 32);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(13, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "c";
            // 
            // groupBoxSelectObjects
            // 
            this.groupBoxSelectObjects.Controls.Add(this.radioButtonChiefs);
            this.groupBoxSelectObjects.Controls.Add(this.comboBoxAuthor);
            this.groupBoxSelectObjects.Controls.Add(this.comboBoxDepartment);
            this.groupBoxSelectObjects.Controls.Add(this.radioButtonObjectsByAuthor);
            this.groupBoxSelectObjects.Controls.Add(this.radioButtonObjectsByDepartment);
            this.groupBoxSelectObjects.Controls.Add(this.radioButtonAllObjects);
            this.groupBoxSelectObjects.Location = new System.Drawing.Point(13, 178);
            this.groupBoxSelectObjects.Name = "groupBoxSelectObjects";
            this.groupBoxSelectObjects.Size = new System.Drawing.Size(489, 92);
            this.groupBoxSelectObjects.TabIndex = 7;
            this.groupBoxSelectObjects.TabStop = false;
            this.groupBoxSelectObjects.Text = "Выбор объектов";
            // 
            // comboBoxAuthor
            // 
            this.comboBoxAuthor.Enabled = false;
            this.comboBoxAuthor.FormattingEnabled = true;
            this.comboBoxAuthor.Location = new System.Drawing.Point(215, 50);
            this.comboBoxAuthor.Name = "comboBoxAuthor";
            this.comboBoxAuthor.Size = new System.Drawing.Size(237, 21);
            this.comboBoxAuthor.TabIndex = 4;
            // 
            // comboBoxDepartment
            // 
            this.comboBoxDepartment.Enabled = false;
            this.comboBoxDepartment.FormattingEnabled = true;
            this.comboBoxDepartment.Location = new System.Drawing.Point(215, 27);
            this.comboBoxDepartment.Name = "comboBoxDepartment";
            this.comboBoxDepartment.Size = new System.Drawing.Size(237, 21);
            this.comboBoxDepartment.TabIndex = 3;
            // 
            // radioButtonObjectsByAuthor
            // 
            this.radioButtonObjectsByAuthor.AutoSize = true;
            this.radioButtonObjectsByAuthor.Location = new System.Drawing.Point(40, 51);
            this.radioButtonObjectsByAuthor.Name = "radioButtonObjectsByAuthor";
            this.radioButtonObjectsByAuthor.Size = new System.Drawing.Size(123, 17);
            this.radioButtonObjectsByAuthor.TabIndex = 2;
            this.radioButtonObjectsByAuthor.Text = "Объекты по автору";
            this.radioButtonObjectsByAuthor.UseVisualStyleBackColor = true;
            this.radioButtonObjectsByAuthor.CheckedChanged += new System.EventHandler(this.radioButtonMyObjects_CheckedChanged);
            // 
            // radioButtonObjectsByDepartment
            // 
            this.radioButtonObjectsByDepartment.AutoSize = true;
            this.radioButtonObjectsByDepartment.Location = new System.Drawing.Point(40, 32);
            this.radioButtonObjectsByDepartment.Name = "radioButtonObjectsByDepartment";
            this.radioButtonObjectsByDepartment.Size = new System.Drawing.Size(169, 17);
            this.radioButtonObjectsByDepartment.TabIndex = 1;
            this.radioButtonObjectsByDepartment.Text = "Объекты по подразделению";
            this.radioButtonObjectsByDepartment.UseVisualStyleBackColor = true;
            this.radioButtonObjectsByDepartment.CheckedChanged += new System.EventHandler(this.radioButtonObjectsMyDepartment_CheckedChanged);
            // 
            // radioButtonAllObjects
            // 
            this.radioButtonAllObjects.AutoSize = true;
            this.radioButtonAllObjects.Checked = true;
            this.radioButtonAllObjects.Location = new System.Drawing.Point(40, 14);
            this.radioButtonAllObjects.Name = "radioButtonAllObjects";
            this.radioButtonAllObjects.Size = new System.Drawing.Size(91, 17);
            this.radioButtonAllObjects.TabIndex = 0;
            this.radioButtonAllObjects.TabStop = true;
            this.radioButtonAllObjects.Text = "Все объекты";
            this.radioButtonAllObjects.UseVisualStyleBackColor = true;
            this.radioButtonAllObjects.CheckedChanged += new System.EventHandler(this.radioButtonAllObjects_CheckedChanged);
            // 
            // checkBoxIsState
            // 
            this.checkBoxIsState.AutoSize = true;
            this.checkBoxIsState.Location = new System.Drawing.Point(19, 22);
            this.checkBoxIsState.Name = "checkBoxIsState";
            this.checkBoxIsState.Size = new System.Drawing.Size(402, 17);
            this.checkBoxIsState.TabIndex = 8;
            this.checkBoxIsState.Text = "Выводить только этапы с превышением срока согласования документов";
            this.checkBoxIsState.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.NumericUDCountControlDays);
            this.groupBox1.Controls.Add(this.checkBoxIsState);
            this.groupBox1.Location = new System.Drawing.Point(11, 83);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(489, 87);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Контроль нормативного срока выполнения работы";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(261, 68);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(212, 13);
            this.label7.TabIndex = 12;
            this.label7.Text = "поступления документа в подраделение";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(319, 47);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(109, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "следующий за днем";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 47);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(242, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Нормативный срок согласования документов";
            // 
            // NumericUDCountControlDays
            // 
            this.NumericUDCountControlDays.Location = new System.Drawing.Point(264, 45);
            this.NumericUDCountControlDays.Name = "NumericUDCountControlDays";
            this.NumericUDCountControlDays.Size = new System.Drawing.Size(49, 20);
            this.NumericUDCountControlDays.TabIndex = 9;
            this.NumericUDCountControlDays.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // radioButtonChiefs
            // 
            this.radioButtonChiefs.AutoSize = true;
            this.radioButtonChiefs.Location = new System.Drawing.Point(40, 71);
            this.radioButtonChiefs.Name = "radioButtonChiefs";
            this.radioButtonChiefs.Size = new System.Drawing.Size(196, 17);
            this.radioButtonChiefs.TabIndex = 5;
            this.radioButtonChiefs.TabStop = true;
            this.radioButtonChiefs.Text = "Объекты только у руководителей";
            this.radioButtonChiefs.UseVisualStyleBackColor = true;
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(520, 318);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBoxSelectObjects);
            this.Controls.Add(this.groupBoxSelectPeriod);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.comboBoxState);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxProcess);
            this.Controls.Add(this.buttonMakeReport);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройка отчета \"Текущие состояния процесса\"";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBoxSelectPeriod.ResumeLayout(false);
            this.groupBoxSelectPeriod.PerformLayout();
            this.groupBoxSelectObjects.ResumeLayout(false);
            this.groupBoxSelectObjects.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NumericUDCountControlDays)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.ComboBox comboBoxProcess;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox comboBoxState;
        private System.Windows.Forms.GroupBox groupBoxSelectPeriod;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dateTimePickerBegin;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBoxSelectObjects;
        private System.Windows.Forms.RadioButton radioButtonObjectsByAuthor;
        private System.Windows.Forms.RadioButton radioButtonObjectsByDepartment;
        private System.Windows.Forms.RadioButton radioButtonAllObjects;
        private System.Windows.Forms.ComboBox comboBoxAuthor;
        private System.Windows.Forms.ComboBox comboBoxDepartment;
        private System.Windows.Forms.CheckBox checkBoxIsState;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown NumericUDCountControlDays;
        private System.Windows.Forms.RadioButton radioButtonChiefs;
    }
}
