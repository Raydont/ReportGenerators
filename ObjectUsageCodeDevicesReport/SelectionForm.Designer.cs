using System;

namespace ObjectUsageCodeDevicesReport
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
            this.SearchDevicesButton = new System.Windows.Forms.Button();
            this.textBoxListNames = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rBOnlyCodeDevicesName = new System.Windows.Forms.RadioButton();
            this.rBFullInformation = new System.Windows.Forms.RadioButton();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // SearchDevicesButton
            // 
            this.SearchDevicesButton.Location = new System.Drawing.Point(23, 450);
            this.SearchDevicesButton.Name = "SearchDevicesButton";
            this.SearchDevicesButton.Size = new System.Drawing.Size(347, 23);
            this.SearchDevicesButton.TabIndex = 0;
            this.SearchDevicesButton.Text = "Сформировать отчет";
            this.SearchDevicesButton.UseVisualStyleBackColor = true;
            this.SearchDevicesButton.Click += new System.EventHandler(this.SearchDevicesButton_Click);
            // 
            // textBoxListNames
            // 
            this.textBoxListNames.Location = new System.Drawing.Point(23, 35);
            this.textBoxListNames.Multiline = true;
            this.textBoxListNames.Name = "textBoxListNames";
            this.textBoxListNames.Size = new System.Drawing.Size(444, 310);
            this.textBoxListNames.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(447, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Список наименований ПКИ (каждое ПКИ с новой строки без разделительных знаков)";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rBOnlyCodeDevicesName);
            this.groupBox1.Controls.Add(this.rBFullInformation);
            this.groupBox1.Location = new System.Drawing.Point(23, 362);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(444, 72);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Тип отчета";
            // 
            // rBOnlyCodeDevicesName
            // 
            this.rBOnlyCodeDevicesName.AutoSize = true;
            this.rBOnlyCodeDevicesName.Checked = true;
            this.rBOnlyCodeDevicesName.Location = new System.Drawing.Point(93, 19);
            this.rBOnlyCodeDevicesName.Name = "rBOnlyCodeDevicesName";
            this.rBOnlyCodeDevicesName.Size = new System.Drawing.Size(197, 17);
            this.rBOnlyCodeDevicesName.TabIndex = 1;
            this.rBOnlyCodeDevicesName.TabStop = true;
            this.rBOnlyCodeDevicesName.Text = "Только шифрованные устройства";
            this.rBOnlyCodeDevicesName.UseVisualStyleBackColor = true;
            // 
            // rBFullInformation
            // 
            this.rBFullInformation.AutoSize = true;
            this.rBFullInformation.Location = new System.Drawing.Point(93, 42);
            this.rBFullInformation.Name = "rBFullInformation";
            this.rBFullInformation.Size = new System.Drawing.Size(290, 17);
            this.rBFullInformation.TabIndex = 0;
            this.rBFullInformation.Text = "Полная информация по всем уровням вложенности";
            this.rBFullInformation.UseVisualStyleBackColor = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(376, 450);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(91, 23);
            this.buttonCancel.TabIndex = 4;
            this.buttonCancel.Text = "Отмена";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 488);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxListNames);
            this.Controls.Add(this.SearchDevicesButton);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Формирование отчета";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SearchDevicesButton;
        private System.Windows.Forms.TextBox textBoxListNames;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rBOnlyCodeDevicesName;
        private System.Windows.Forms.RadioButton rBFullInformation;
        private System.Windows.Forms.Button buttonCancel;
    }
}