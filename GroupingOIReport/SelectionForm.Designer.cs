namespace GroupingOIReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SelectionForm));
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.buttonLoadFile1 = new System.Windows.Forms.Button();
            this.textBoxFileName1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBoxOtherItems = new System.Windows.Forms.CheckBox();
            this.checkBoxStandartItems = new System.Windows.Forms.CheckBox();
            this.groupBoxTypes = new System.Windows.Forms.GroupBox();
            this.checkBoxDetails = new System.Windows.Forms.CheckBox();
            this.groupBoxOptions = new System.Windows.Forms.GroupBox();
            this.checkBoxIsCreateObjectCost = new System.Windows.Forms.CheckBox();
            this.checkBoxIsResultManyFiles = new System.Windows.Forms.CheckBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.checkBoxNotCodeDevices = new System.Windows.Forms.CheckBox();
            this.checkBoxAddAllCodeObject = new System.Windows.Forms.CheckBox();
            this.checkBoxAddCodeDevices = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxOrderNumber = new System.Windows.Forms.TextBox();
            this.textBoxSpecificationOrder = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.groupBoxTypes.SuspendLayout();
            this.groupBoxOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Enabled = false;
            this.buttonMakeReport.Location = new System.Drawing.Point(21, 421);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(536, 23);
            this.buttonMakeReport.TabIndex = 0;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // buttonLoadFile1
            // 
            this.buttonLoadFile1.Location = new System.Drawing.Point(522, 16);
            this.buttonLoadFile1.Name = "buttonLoadFile1";
            this.buttonLoadFile1.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadFile1.TabIndex = 6;
            this.buttonLoadFile1.Text = "...";
            this.buttonLoadFile1.UseVisualStyleBackColor = true;
            this.buttonLoadFile1.Click += new System.EventHandler(this.buttonLoadFile1_Click);
            // 
            // textBoxFileName1
            // 
            this.textBoxFileName1.Location = new System.Drawing.Point(131, 17);
            this.textBoxFileName1.Name = "textBoxFileName1";
            this.textBoxFileName1.Size = new System.Drawing.Size(401, 20);
            this.textBoxFileName1.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Загружаемый файл";
            // 
            // checkBoxOtherItems
            // 
            this.checkBoxOtherItems.AutoSize = true;
            this.checkBoxOtherItems.Checked = true;
            this.checkBoxOtherItems.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxOtherItems.Location = new System.Drawing.Point(40, 23);
            this.checkBoxOtherItems.Name = "checkBoxOtherItems";
            this.checkBoxOtherItems.Size = new System.Drawing.Size(108, 17);
            this.checkBoxOtherItems.TabIndex = 8;
            this.checkBoxOtherItems.Text = "Прочие изделия";
            this.checkBoxOtherItems.UseVisualStyleBackColor = true;
            // 
            // checkBoxStandartItems
            // 
            this.checkBoxStandartItems.AutoSize = true;
            this.checkBoxStandartItems.Checked = true;
            this.checkBoxStandartItems.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxStandartItems.Location = new System.Drawing.Point(342, 23);
            this.checkBoxStandartItems.Name = "checkBoxStandartItems";
            this.checkBoxStandartItems.Size = new System.Drawing.Size(138, 17);
            this.checkBoxStandartItems.TabIndex = 9;
            this.checkBoxStandartItems.Text = "Стандартные изделия";
            this.checkBoxStandartItems.UseVisualStyleBackColor = true;
            // 
            // groupBoxTypes
            // 
            this.groupBoxTypes.Controls.Add(this.checkBoxDetails);
            this.groupBoxTypes.Controls.Add(this.checkBoxStandartItems);
            this.groupBoxTypes.Controls.Add(this.checkBoxOtherItems);
            this.groupBoxTypes.Location = new System.Drawing.Point(17, 267);
            this.groupBoxTypes.Name = "groupBoxTypes";
            this.groupBoxTypes.Size = new System.Drawing.Size(512, 54);
            this.groupBoxTypes.TabIndex = 11;
            this.groupBoxTypes.TabStop = false;
            this.groupBoxTypes.Text = "Типы объектов";
            // 
            // checkBoxDetails
            // 
            this.checkBoxDetails.AutoSize = true;
            this.checkBoxDetails.Checked = true;
            this.checkBoxDetails.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxDetails.Location = new System.Drawing.Point(211, 23);
            this.checkBoxDetails.Name = "checkBoxDetails";
            this.checkBoxDetails.Size = new System.Drawing.Size(64, 17);
            this.checkBoxDetails.TabIndex = 10;
            this.checkBoxDetails.Text = "Детали";
            this.checkBoxDetails.UseVisualStyleBackColor = true;
            // 
            // groupBoxOptions
            // 
            this.groupBoxOptions.Controls.Add(this.checkBoxIsCreateObjectCost);
            this.groupBoxOptions.Controls.Add(this.checkBoxIsResultManyFiles);
            this.groupBoxOptions.Controls.Add(this.textBox1);
            this.groupBoxOptions.Controls.Add(this.checkBoxNotCodeDevices);
            this.groupBoxOptions.Controls.Add(this.checkBoxAddAllCodeObject);
            this.groupBoxOptions.Controls.Add(this.checkBoxAddCodeDevices);
            this.groupBoxOptions.Controls.Add(this.label3);
            this.groupBoxOptions.Controls.Add(this.label1);
            this.groupBoxOptions.Controls.Add(this.textBoxOrderNumber);
            this.groupBoxOptions.Controls.Add(this.textBoxSpecificationOrder);
            this.groupBoxOptions.Controls.Add(this.groupBoxTypes);
            this.groupBoxOptions.Location = new System.Drawing.Point(21, 51);
            this.groupBoxOptions.Name = "groupBoxOptions";
            this.groupBoxOptions.Size = new System.Drawing.Size(536, 357);
            this.groupBoxOptions.TabIndex = 12;
            this.groupBoxOptions.TabStop = false;
            this.groupBoxOptions.Text = "Настройки отчета";
            // 
            // checkBoxIsCreateObjectCost
            // 
            this.checkBoxIsCreateObjectCost.AutoSize = true;
            this.checkBoxIsCreateObjectCost.Location = new System.Drawing.Point(274, 327);
            this.checkBoxIsCreateObjectCost.Name = "checkBoxIsCreateObjectCost";
            this.checkBoxIsCreateObjectCost.Size = new System.Drawing.Size(159, 17);
            this.checkBoxIsCreateObjectCost.TabIndex = 22;
            this.checkBoxIsCreateObjectCost.Text = "Выводить стоимость ПКИ";
            this.checkBoxIsCreateObjectCost.UseVisualStyleBackColor = true;
            // 
            // checkBoxIsResultManyFiles
            // 
            this.checkBoxIsResultManyFiles.AutoSize = true;
            this.checkBoxIsResultManyFiles.Location = new System.Drawing.Point(17, 327);
            this.checkBoxIsResultManyFiles.Name = "checkBoxIsResultManyFiles";
            this.checkBoxIsResultManyFiles.Size = new System.Drawing.Size(237, 17);
            this.checkBoxIsResultManyFiles.TabIndex = 21;
            this.checkBoxIsResultManyFiles.Text = "Выводить результат в отдельные файлы ";
            this.checkBoxIsResultManyFiles.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.textBox1.Location = new System.Drawing.Point(17, 178);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(512, 77);
            this.textBox1.TabIndex = 20;
            this.textBox1.Text = resources.GetString("textBox1.Text");
            // 
            // checkBoxNotCodeDevices
            // 
            this.checkBoxNotCodeDevices.AutoSize = true;
            this.checkBoxNotCodeDevices.Location = new System.Drawing.Point(18, 152);
            this.checkBoxNotCodeDevices.Name = "checkBoxNotCodeDevices";
            this.checkBoxNotCodeDevices.Size = new System.Drawing.Size(511, 17);
            this.checkBoxNotCodeDevices.TabIndex = 18;
            this.checkBoxNotCodeDevices.Text = "Выводить только состав объектов,  указанных в файле без входящих шифрованных устр" +
    "ойств";
            this.checkBoxNotCodeDevices.UseVisualStyleBackColor = true;
            this.checkBoxNotCodeDevices.CheckedChanged += new System.EventHandler(this.checkBoxFirstLevel_CheckedChanged);
            // 
            // checkBoxAddAllCodeObject
            // 
            this.checkBoxAddAllCodeObject.AutoSize = true;
            this.checkBoxAddAllCodeObject.Location = new System.Drawing.Point(18, 125);
            this.checkBoxAddAllCodeObject.Name = "checkBoxAddAllCodeObject";
            this.checkBoxAddAllCodeObject.Size = new System.Drawing.Size(445, 17);
            this.checkBoxAddAllCodeObject.TabIndex = 17;
            this.checkBoxAddAllCodeObject.Text = "Добавлять все найденные шифрованные устройства (плюс к указанным в файле)";
            this.checkBoxAddAllCodeObject.UseVisualStyleBackColor = true;
            this.checkBoxAddAllCodeObject.CheckedChanged += new System.EventHandler(this.checkBoxAddAllCodeObject_CheckedChanged);
            // 
            // checkBoxAddCodeDevices
            // 
            this.checkBoxAddCodeDevices.AutoSize = true;
            this.checkBoxAddCodeDevices.Checked = true;
            this.checkBoxAddCodeDevices.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxAddCodeDevices.Location = new System.Drawing.Point(18, 97);
            this.checkBoxAddCodeDevices.Name = "checkBoxAddCodeDevices";
            this.checkBoxAddCodeDevices.Size = new System.Drawing.Size(380, 17);
            this.checkBoxAddCodeDevices.TabIndex = 16;
            this.checkBoxAddCodeDevices.Text = "Добавлять отдельно шифрованные устройства из состава устройств";
            this.checkBoxAddCodeDevices.UseVisualStyleBackColor = true;
            this.checkBoxAddCodeDevices.CheckedChanged += new System.EventHandler(this.checkBoxAddCodeDevices_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "Номер наряд-заказа";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 14;
            this.label1.Text = "Номер заказа";
            // 
            // textBoxOrderNumber
            // 
            this.textBoxOrderNumber.Location = new System.Drawing.Point(139, 28);
            this.textBoxOrderNumber.Name = "textBoxOrderNumber";
            this.textBoxOrderNumber.Size = new System.Drawing.Size(372, 20);
            this.textBoxOrderNumber.TabIndex = 13;
            // 
            // textBoxSpecificationOrder
            // 
            this.textBoxSpecificationOrder.Location = new System.Drawing.Point(139, 59);
            this.textBoxSpecificationOrder.Name = "textBoxSpecificationOrder";
            this.textBoxSpecificationOrder.Size = new System.Drawing.Size(372, 20);
            this.textBoxSpecificationOrder.TabIndex = 12;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(569, 456);
            this.Controls.Add(this.groupBoxOptions);
            this.Controls.Add(this.buttonLoadFile1);
            this.Controls.Add(this.textBoxFileName1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonMakeReport);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройка отчета";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBoxTypes.ResumeLayout(false);
            this.groupBoxTypes.PerformLayout();
            this.groupBoxOptions.ResumeLayout(false);
            this.groupBoxOptions.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.Button buttonLoadFile1;
        private System.Windows.Forms.TextBox textBoxFileName1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBoxOtherItems;
        private System.Windows.Forms.CheckBox checkBoxStandartItems;
        private System.Windows.Forms.GroupBox groupBoxTypes;
        private System.Windows.Forms.GroupBox groupBoxOptions;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxOrderNumber;
        private System.Windows.Forms.TextBox textBoxSpecificationOrder;
        private System.Windows.Forms.CheckBox checkBoxAddCodeDevices;
        private System.Windows.Forms.CheckBox checkBoxAddAllCodeObject;
        private System.Windows.Forms.CheckBox checkBoxNotCodeDevices;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.CheckBox checkBoxIsResultManyFiles;
        private System.Windows.Forms.CheckBox checkBoxIsCreateObjectCost;
        private System.Windows.Forms.CheckBox checkBoxDetails;
    }
}