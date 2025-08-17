namespace CalculationOtherProducts
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkMaterial = new System.Windows.Forms.CheckBox();
            this.chkOther = new System.Windows.Forms.CheckBox();
            this.chkStandart = new System.Windows.Forms.CheckBox();
            this.chkAssembly = new System.Windows.Forms.CheckBox();
            this.chkDetail = new System.Windows.Forms.CheckBox();
            this.btnMakeReport = new System.Windows.Forms.Button();
            this.buttonLoadListDevice = new System.Windows.Forms.Button();
            this.textBoxListDevice = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonLoadListPKI = new System.Windows.Forms.Button();
            this.textBoxListPKI = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.openFileDialogListDevice = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialogListPKI = new System.Windows.Forms.OpenFileDialog();
            this.checkBoxUseCurrentObject = new System.Windows.Forms.CheckBox();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkMaterial);
            this.groupBox2.Controls.Add(this.chkOther);
            this.groupBox2.Controls.Add(this.chkStandart);
            this.groupBox2.Controls.Add(this.chkAssembly);
            this.groupBox2.Controls.Add(this.chkDetail);
            this.groupBox2.Location = new System.Drawing.Point(12, 121);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(436, 89);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Типы объектов";
            // 
            // chkMaterial
            // 
            this.chkMaterial.AutoSize = true;
            this.chkMaterial.Location = new System.Drawing.Point(26, 51);
            this.chkMaterial.Name = "chkMaterial";
            this.chkMaterial.Size = new System.Drawing.Size(84, 17);
            this.chkMaterial.TabIndex = 6;
            this.chkMaterial.Text = "Материалы";
            this.chkMaterial.UseVisualStyleBackColor = true;
            // 
            // chkOther
            // 
            this.chkOther.AutoSize = true;
            this.chkOther.Checked = true;
            this.chkOther.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOther.Location = new System.Drawing.Point(26, 28);
            this.chkOther.Name = "chkOther";
            this.chkOther.Size = new System.Drawing.Size(108, 17);
            this.chkOther.TabIndex = 5;
            this.chkOther.Text = "Прочие изделия";
            this.chkOther.UseVisualStyleBackColor = true;
            // 
            // chkStandart
            // 
            this.chkStandart.AutoSize = true;
            this.chkStandart.Location = new System.Drawing.Point(173, 51);
            this.chkStandart.Name = "chkStandart";
            this.chkStandart.Size = new System.Drawing.Size(138, 17);
            this.chkStandart.TabIndex = 4;
            this.chkStandart.Text = "Стандартные изделия";
            this.chkStandart.UseVisualStyleBackColor = true;
            // 
            // chkAssembly
            // 
            this.chkAssembly.AutoSize = true;
            this.chkAssembly.Location = new System.Drawing.Point(173, 28);
            this.chkAssembly.Name = "chkAssembly";
            this.chkAssembly.Size = new System.Drawing.Size(129, 17);
            this.chkAssembly.TabIndex = 3;
            this.chkAssembly.Text = "Сборочные единицы";
            this.chkAssembly.UseVisualStyleBackColor = true;
            // 
            // chkDetail
            // 
            this.chkDetail.AutoSize = true;
            this.chkDetail.Location = new System.Drawing.Point(360, 28);
            this.chkDetail.Name = "chkDetail";
            this.chkDetail.Size = new System.Drawing.Size(64, 17);
            this.chkDetail.TabIndex = 2;
            this.chkDetail.Text = "Детали";
            this.chkDetail.UseVisualStyleBackColor = true;
            // 
            // btnMakeReport
            // 
            this.btnMakeReport.Location = new System.Drawing.Point(12, 259);
            this.btnMakeReport.Name = "btnMakeReport";
            this.btnMakeReport.Size = new System.Drawing.Size(436, 23);
            this.btnMakeReport.TabIndex = 5;
            this.btnMakeReport.Text = "Сформировать отчет";
            this.btnMakeReport.UseVisualStyleBackColor = true;
            this.btnMakeReport.Click += new System.EventHandler(this.btnMakeReport_Click);
            // 
            // buttonLoadListDevice
            // 
            this.buttonLoadListDevice.Location = new System.Drawing.Point(389, 25);
            this.buttonLoadListDevice.Name = "buttonLoadListDevice";
            this.buttonLoadListDevice.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadListDevice.TabIndex = 9;
            this.buttonLoadListDevice.Text = "...";
            this.buttonLoadListDevice.UseVisualStyleBackColor = true;
            this.buttonLoadListDevice.Click += new System.EventHandler(this.buttonLoadListDevice_Click);
            // 
            // textBoxListDevice
            // 
            this.textBoxListDevice.Location = new System.Drawing.Point(110, 27);
            this.textBoxListDevice.Name = "textBoxListDevice";
            this.textBoxListDevice.Size = new System.Drawing.Size(303, 20);
            this.textBoxListDevice.TabIndex = 8;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Список устройств";
            // 
            // buttonLoadListPKI
            // 
            this.buttonLoadListPKI.Location = new System.Drawing.Point(390, 55);
            this.buttonLoadListPKI.Name = "buttonLoadListPKI";
            this.buttonLoadListPKI.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadListPKI.TabIndex = 12;
            this.buttonLoadListPKI.Text = "...";
            this.buttonLoadListPKI.UseVisualStyleBackColor = true;
            this.buttonLoadListPKI.Click += new System.EventHandler(this.buttonLoadListPKI_Click);
            // 
            // textBoxListPKI
            // 
            this.textBoxListPKI.Location = new System.Drawing.Point(110, 57);
            this.textBoxListPKI.Name = "textBoxListPKI";
            this.textBoxListPKI.Size = new System.Drawing.Size(303, 20);
            this.textBoxListPKI.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Список ПКИ";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.buttonLoadListDevice);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.buttonLoadListPKI);
            this.groupBox1.Controls.Add(this.textBoxListDevice);
            this.groupBox1.Controls.Add(this.textBoxListPKI);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 13);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(436, 95);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Загружаемые файлы";
            // 
            // checkBoxUseCurrentObject
            // 
            this.checkBoxUseCurrentObject.AutoSize = true;
            this.checkBoxUseCurrentObject.Checked = true;
            this.checkBoxUseCurrentObject.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxUseCurrentObject.Location = new System.Drawing.Point(12, 228);
            this.checkBoxUseCurrentObject.Name = "checkBoxUseCurrentObject";
            this.checkBoxUseCurrentObject.Size = new System.Drawing.Size(217, 17);
            this.checkBoxUseCurrentObject.TabIndex = 15;
            this.checkBoxUseCurrentObject.Text = "Разузловывавть выделенный объект";
            this.checkBoxUseCurrentObject.UseVisualStyleBackColor = true;
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(460, 294);
            this.Controls.Add(this.checkBoxUseCurrentObject);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnMakeReport);
            this.Controls.Add(this.groupBox2);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Вывод структуры и группировка объектов";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkMaterial;
        private System.Windows.Forms.CheckBox chkOther;
        private System.Windows.Forms.CheckBox chkStandart;
        private System.Windows.Forms.CheckBox chkAssembly;
        private System.Windows.Forms.CheckBox chkDetail;
        private System.Windows.Forms.Button btnMakeReport;
        private System.Windows.Forms.Button buttonLoadListDevice;
        private System.Windows.Forms.TextBox textBoxListDevice;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonLoadListPKI;
        private System.Windows.Forms.TextBox textBoxListPKI;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialogListDevice;
        private System.Windows.Forms.OpenFileDialog openFileDialogListPKI;
        private System.Windows.Forms.CheckBox checkBoxUseCurrentObject;
    }
}