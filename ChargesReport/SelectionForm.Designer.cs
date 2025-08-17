using System;
using TFlex.Reporting;

namespace ChargesReport
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
            this.comboBoxMonth = new System.Windows.Forms.ComboBox();
            this.comboBoxYear = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBoxMethodPayment = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBoxDepartment = new System.Windows.Forms.ComboBox();
            this.textBoxCountWorkHours = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBoxTypeReport = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // makeReportButton
            // 
            this.makeReportButton.Location = new System.Drawing.Point(28, 206);
            this.makeReportButton.Name = "makeReportButton";
            this.makeReportButton.Size = new System.Drawing.Size(294, 23);
            this.makeReportButton.TabIndex = 7;
            this.makeReportButton.Text = "Создать отчет";
            this.makeReportButton.UseVisualStyleBackColor = true;
            this.makeReportButton.Click += new System.EventHandler(this.makeReportButton_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Файлы IPC|*.ipc";
            // 
            // comboBoxMonth
            // 
            this.comboBoxMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
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
            this.comboBoxMonth.Location = new System.Drawing.Point(129, 27);
            this.comboBoxMonth.Name = "comboBoxMonth";
            this.comboBoxMonth.Size = new System.Drawing.Size(193, 21);
            this.comboBoxMonth.TabIndex = 9;
            this.comboBoxMonth.SelectedIndexChanged += new System.EventHandler(this.comboBoxMonth_SelectedIndexChanged);
            // 
            // comboBoxYear
            // 
            this.comboBoxYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxYear.FormattingEnabled = true;
            this.comboBoxYear.Location = new System.Drawing.Point(129, 57);
            this.comboBoxYear.Name = "comboBoxYear";
            this.comboBoxYear.Size = new System.Drawing.Size(193, 21);
            this.comboBoxYear.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Выбор месяца";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(66, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Выбор года";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(25, 120);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Способ оплаты";
            // 
            // comboBoxMethodPayment
            // 
            this.comboBoxMethodPayment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxMethodPayment.FormattingEnabled = true;
            this.comboBoxMethodPayment.Items.AddRange(new object[] {
            "сдельная",
            "повременная"});
            this.comboBoxMethodPayment.Location = new System.Drawing.Point(129, 117);
            this.comboBoxMethodPayment.Name = "comboBoxMethodPayment";
            this.comboBoxMethodPayment.Size = new System.Drawing.Size(193, 21);
            this.comboBoxMethodPayment.TabIndex = 13;
            this.comboBoxMethodPayment.SelectedIndexChanged += new System.EventHandler(this.comboBoxMethodPayment_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(25, 90);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Подразделение";
            // 
            // comboBoxDepartment
            // 
            this.comboBoxDepartment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxDepartment.FormattingEnabled = true;
            this.comboBoxDepartment.Items.AddRange(new object[] {
            "ПДО",
            "РСУ 966",
            "Цех 75",
            "Цех 80",
            "Цех 85",
            "Участок 614",
            "Цех 70"});
            this.comboBoxDepartment.Location = new System.Drawing.Point(129, 87);
            this.comboBoxDepartment.Name = "comboBoxDepartment";
            this.comboBoxDepartment.Size = new System.Drawing.Size(193, 21);
            this.comboBoxDepartment.TabIndex = 15;
            // 
            // textBoxCountWorkHours
            // 
            this.textBoxCountWorkHours.Location = new System.Drawing.Point(129, 177);
            this.textBoxCountWorkHours.Name = "textBoxCountWorkHours";
            this.textBoxCountWorkHours.Size = new System.Drawing.Size(193, 20);
            this.textBoxCountWorkHours.TabIndex = 17;
            this.textBoxCountWorkHours.Text = "176,5";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(25, 180);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(97, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "Кол-во раб. часов";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(25, 150);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(62, 13);
            this.label6.TabIndex = 20;
            this.label6.Text = "Вид отчета";
            // 
            // comboBoxTypeReport
            // 
            this.comboBoxTypeReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxTypeReport.Enabled = false;
            this.comboBoxTypeReport.FormattingEnabled = true;
            this.comboBoxTypeReport.Items.AddRange(new object[] {
            "По табельному времени",
            "Выходные - РВ1",
            "Выходные - РВ2",
            "Доплаты",
            "Сверхурочные"});
            this.comboBoxTypeReport.Location = new System.Drawing.Point(129, 147);
            this.comboBoxTypeReport.Name = "comboBoxTypeReport";
            this.comboBoxTypeReport.Size = new System.Drawing.Size(193, 21);
            this.comboBoxTypeReport.TabIndex = 19;
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(349, 251);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBoxTypeReport);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBoxCountWorkHours);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.comboBoxDepartment);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBoxMethodPayment);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxYear);
            this.Controls.Add(this.comboBoxMonth);
            this.Controls.Add(this.makeReportButton);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ведомость начисления премии рабочим";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button makeReportButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.ComboBox comboBoxMonth;
        private System.Windows.Forms.ComboBox comboBoxYear;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox comboBoxMethodPayment;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBoxDepartment;
        private System.Windows.Forms.TextBox textBoxCountWorkHours;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBoxTypeReport;
    }
}