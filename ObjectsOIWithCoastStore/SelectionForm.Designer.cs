namespace ObjectsOIWithCoastStore
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
            this.buttonLoadFileListPKI = new System.Windows.Forms.Button();
            this.textBoxFileNameListPKI = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.radioButtonForSelectedObject = new System.Windows.Forms.RadioButton();
            this.radioButtonForListPKI = new System.Windows.Forms.RadioButton();
            this.radioButtonForListDevice = new System.Windows.Forms.RadioButton();
            this.buttonLoadFileListDevices = new System.Windows.Forms.Button();
            this.textBoxFileNameListDevice = new System.Windows.Forms.TextBox();
            this.radioButtonForListNomenclature = new System.Windows.Forms.RadioButton();
            this.buttonLoadFileListNomenclature = new System.Windows.Forms.Button();
            this.textBoxFileNameListNomenclature = new System.Windows.Forms.TextBox();
            this.buttonLoadFileListDevicesKoef = new System.Windows.Forms.Button();
            this.textBoxFileNameDevicesKoef = new System.Windows.Forms.TextBox();
            this.radioButtonForListDeviceKoef = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.buttonLoadListDevicesDifferentFiles = new System.Windows.Forms.Button();
            this.textBoxFileNameListDevicesDifferentFiles = new System.Windows.Forms.TextBox();
            this.radioButtonForListDevicesDifferentFiles = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.radioButtonForOIFailureByPeriod = new System.Windows.Forms.RadioButton();
            this.labelEnd = new System.Windows.Forms.Label();
            this.dateTimePickerBegin = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.labelBegin = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(24, 292);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(547, 23);
            this.buttonMakeReport.TabIndex = 0;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // buttonLoadFileListPKI
            // 
            this.buttonLoadFileListPKI.Enabled = false;
            this.buttonLoadFileListPKI.Location = new System.Drawing.Point(536, 54);
            this.buttonLoadFileListPKI.Name = "buttonLoadFileListPKI";
            this.buttonLoadFileListPKI.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadFileListPKI.TabIndex = 6;
            this.buttonLoadFileListPKI.Text = "...";
            this.buttonLoadFileListPKI.UseVisualStyleBackColor = true;
            this.buttonLoadFileListPKI.Click += new System.EventHandler(this.buttonLoadFile1_Click);
            // 
            // textBoxFileNameListPKI
            // 
            this.textBoxFileNameListPKI.Enabled = false;
            this.textBoxFileNameListPKI.Location = new System.Drawing.Point(224, 56);
            this.textBoxFileNameListPKI.Name = "textBoxFileNameListPKI";
            this.textBoxFileNameListPKI.Size = new System.Drawing.Size(318, 20);
            this.textBoxFileNameListPKI.TabIndex = 5;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // radioButtonForSelectedObject
            // 
            this.radioButtonForSelectedObject.AutoSize = true;
            this.radioButtonForSelectedObject.Checked = true;
            this.radioButtonForSelectedObject.Location = new System.Drawing.Point(34, 27);
            this.radioButtonForSelectedObject.Name = "radioButtonForSelectedObject";
            this.radioButtonForSelectedObject.Size = new System.Drawing.Size(58, 17);
            this.radioButtonForSelectedObject.TabIndex = 8;
            this.radioButtonForSelectedObject.TabStop = true;
            this.radioButtonForSelectedObject.Text = "Для ...";
            this.radioButtonForSelectedObject.UseVisualStyleBackColor = true;
            this.radioButtonForSelectedObject.CheckedChanged += new System.EventHandler(this.radioButtonForSelectedObject_CheckedChanged);
            // 
            // radioButtonForListPKI
            // 
            this.radioButtonForListPKI.AutoSize = true;
            this.radioButtonForListPKI.Location = new System.Drawing.Point(34, 58);
            this.radioButtonForListPKI.Name = "radioButtonForListPKI";
            this.radioButtonForListPKI.Size = new System.Drawing.Size(151, 17);
            this.radioButtonForListPKI.TabIndex = 9;
            this.radioButtonForListPKI.Text = "Для списка ПКИ (Парус)";
            this.radioButtonForListPKI.UseVisualStyleBackColor = true;
            this.radioButtonForListPKI.CheckedChanged += new System.EventHandler(this.radioButtonForListPKI_CheckedChanged);
            // 
            // radioButtonForListDevice
            // 
            this.radioButtonForListDevice.AutoSize = true;
            this.radioButtonForListDevice.Location = new System.Drawing.Point(34, 120);
            this.radioButtonForListDevice.Name = "radioButtonForListDevice";
            this.radioButtonForListDevice.Size = new System.Drawing.Size(139, 17);
            this.radioButtonForListDevice.TabIndex = 10;
            this.radioButtonForListDevice.Text = "Для списка устройств";
            this.radioButtonForListDevice.UseVisualStyleBackColor = true;
            this.radioButtonForListDevice.CheckedChanged += new System.EventHandler(this.radioButtonForListDevice_CheckedChanged);
            // 
            // buttonLoadFileListDevices
            // 
            this.buttonLoadFileListDevices.Enabled = false;
            this.buttonLoadFileListDevices.Location = new System.Drawing.Point(536, 118);
            this.buttonLoadFileListDevices.Name = "buttonLoadFileListDevices";
            this.buttonLoadFileListDevices.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadFileListDevices.TabIndex = 12;
            this.buttonLoadFileListDevices.Text = "...";
            this.buttonLoadFileListDevices.UseVisualStyleBackColor = true;
            this.buttonLoadFileListDevices.Click += new System.EventHandler(this.buttonLoadFileListDevices_Click);
            // 
            // textBoxFileNameListDevice
            // 
            this.textBoxFileNameListDevice.Enabled = false;
            this.textBoxFileNameListDevice.Location = new System.Drawing.Point(224, 120);
            this.textBoxFileNameListDevice.Name = "textBoxFileNameListDevice";
            this.textBoxFileNameListDevice.Size = new System.Drawing.Size(318, 20);
            this.textBoxFileNameListDevice.TabIndex = 11;
            // 
            // radioButtonForListNomenclature
            // 
            this.radioButtonForListNomenclature.AutoSize = true;
            this.radioButtonForListNomenclature.Location = new System.Drawing.Point(34, 89);
            this.radioButtonForListNomenclature.Name = "radioButtonForListNomenclature";
            this.radioButtonForListNomenclature.Size = new System.Drawing.Size(187, 17);
            this.radioButtonForListNomenclature.TabIndex = 15;
            this.radioButtonForListNomenclature.Text = "Для списка ПКИ (T-FLEX DOCs)";
            this.radioButtonForListNomenclature.UseVisualStyleBackColor = true;
            this.radioButtonForListNomenclature.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // buttonLoadFileListNomenclature
            // 
            this.buttonLoadFileListNomenclature.Enabled = false;
            this.buttonLoadFileListNomenclature.Location = new System.Drawing.Point(536, 85);
            this.buttonLoadFileListNomenclature.Name = "buttonLoadFileListNomenclature";
            this.buttonLoadFileListNomenclature.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadFileListNomenclature.TabIndex = 14;
            this.buttonLoadFileListNomenclature.Text = "...";
            this.buttonLoadFileListNomenclature.UseVisualStyleBackColor = true;
            this.buttonLoadFileListNomenclature.Click += new System.EventHandler(this.buttonLoadFileListNomenclature_Click);
            // 
            // textBoxFileNameListNomenclature
            // 
            this.textBoxFileNameListNomenclature.Enabled = false;
            this.textBoxFileNameListNomenclature.Location = new System.Drawing.Point(224, 87);
            this.textBoxFileNameListNomenclature.Name = "textBoxFileNameListNomenclature";
            this.textBoxFileNameListNomenclature.Size = new System.Drawing.Size(318, 20);
            this.textBoxFileNameListNomenclature.TabIndex = 13;
            // 
            // buttonLoadFileListDevicesKoef
            // 
            this.buttonLoadFileListDevicesKoef.Enabled = false;
            this.buttonLoadFileListDevicesKoef.Location = new System.Drawing.Point(536, 151);
            this.buttonLoadFileListDevicesKoef.Name = "buttonLoadFileListDevicesKoef";
            this.buttonLoadFileListDevicesKoef.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadFileListDevicesKoef.TabIndex = 18;
            this.buttonLoadFileListDevicesKoef.Text = "...";
            this.buttonLoadFileListDevicesKoef.UseVisualStyleBackColor = true;
            this.buttonLoadFileListDevicesKoef.Click += new System.EventHandler(this.buttonLoadFileListDevicesKoef_Click);
            // 
            // textBoxFileNameDevicesKoef
            // 
            this.textBoxFileNameDevicesKoef.Enabled = false;
            this.textBoxFileNameDevicesKoef.Location = new System.Drawing.Point(224, 153);
            this.textBoxFileNameDevicesKoef.Name = "textBoxFileNameDevicesKoef";
            this.textBoxFileNameDevicesKoef.Size = new System.Drawing.Size(318, 20);
            this.textBoxFileNameDevicesKoef.TabIndex = 17;
            // 
            // radioButtonForListDeviceKoef
            // 
            this.radioButtonForListDeviceKoef.AutoSize = true;
            this.radioButtonForListDeviceKoef.Location = new System.Drawing.Point(34, 153);
            this.radioButtonForListDeviceKoef.Name = "radioButtonForListDeviceKoef";
            this.radioButtonForListDeviceKoef.Size = new System.Drawing.Size(148, 17);
            this.radioButtonForListDeviceKoef.TabIndex = 16;
            this.radioButtonForListDeviceKoef.Text = "Для списка устройств с";
            this.radioButtonForListDeviceKoef.UseVisualStyleBackColor = true;
            this.radioButtonForListDeviceKoef.CheckedChanged += new System.EventHandler(this.radioButtonForListDeviceKoef_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(53, 173);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 13);
            this.label1.TabIndex = 19;
            this.label1.Text = "коэффициентами";
            // 
            // statusStrip
            // 
            this.statusStrip.Location = new System.Drawing.Point(0, 332);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(609, 22);
            this.statusStrip.TabIndex = 20;
            this.statusStrip.Text = "statusStrip1";
            // 
            // buttonLoadListDevicesDifferentFiles
            // 
            this.buttonLoadListDevicesDifferentFiles.Enabled = false;
            this.buttonLoadListDevicesDifferentFiles.Location = new System.Drawing.Point(536, 187);
            this.buttonLoadListDevicesDifferentFiles.Name = "buttonLoadListDevicesDifferentFiles";
            this.buttonLoadListDevicesDifferentFiles.Size = new System.Drawing.Size(35, 23);
            this.buttonLoadListDevicesDifferentFiles.TabIndex = 23;
            this.buttonLoadListDevicesDifferentFiles.Text = "...";
            this.buttonLoadListDevicesDifferentFiles.UseVisualStyleBackColor = true;
            this.buttonLoadListDevicesDifferentFiles.Click += new System.EventHandler(this.buttonLoadListDevicesDifferentFiles_Click);
            // 
            // textBoxFileNameListDevicesDifferentFiles
            // 
            this.textBoxFileNameListDevicesDifferentFiles.Enabled = false;
            this.textBoxFileNameListDevicesDifferentFiles.Location = new System.Drawing.Point(224, 189);
            this.textBoxFileNameListDevicesDifferentFiles.Name = "textBoxFileNameListDevicesDifferentFiles";
            this.textBoxFileNameListDevicesDifferentFiles.Size = new System.Drawing.Size(318, 20);
            this.textBoxFileNameListDevicesDifferentFiles.TabIndex = 22;
            // 
            // radioButtonForListDevicesDifferentFiles
            // 
            this.radioButtonForListDevicesDifferentFiles.AutoSize = true;
            this.radioButtonForListDevicesDifferentFiles.Location = new System.Drawing.Point(34, 189);
            this.radioButtonForListDevicesDifferentFiles.Name = "radioButtonForListDevicesDifferentFiles";
            this.radioButtonForListDevicesDifferentFiles.Size = new System.Drawing.Size(139, 17);
            this.radioButtonForListDevicesDifferentFiles.TabIndex = 21;
            this.radioButtonForListDevicesDifferentFiles.Text = "Для списка устройств";
            this.radioButtonForListDevicesDifferentFiles.UseVisualStyleBackColor = true;
            this.radioButtonForListDevicesDifferentFiles.CheckedChanged += new System.EventHandler(this.radioButtonForListDevicesDifferentFiles_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(53, 209);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(153, 13);
            this.label2.TabIndex = 24;
            this.label2.Text = "(обсчет каждого устройства)";
            // 
            // radioButtonForOIFailureByPeriod
            // 
            this.radioButtonForOIFailureByPeriod.AutoSize = true;
            this.radioButtonForOIFailureByPeriod.Location = new System.Drawing.Point(34, 232);
            this.radioButtonForOIFailureByPeriod.Name = "radioButtonForOIFailureByPeriod";
            this.radioButtonForOIFailureByPeriod.Size = new System.Drawing.Size(329, 17);
            this.radioButtonForOIFailureByPeriod.TabIndex = 25;
            this.radioButtonForOIFailureByPeriod.Text = "Для объектов спр-ка Реестр отказов КРЭ и ПКИ за период";
            this.radioButtonForOIFailureByPeriod.UseVisualStyleBackColor = true;
            this.radioButtonForOIFailureByPeriod.CheckedChanged += new System.EventHandler(this.radioButtonForOIFailureByPeriod_CheckedChanged);
            // 
            // labelEnd
            // 
            this.labelEnd.AutoSize = true;
            this.labelEnd.Enabled = false;
            this.labelEnd.Location = new System.Drawing.Point(385, 261);
            this.labelEnd.Name = "labelEnd";
            this.labelEnd.Size = new System.Drawing.Size(19, 13);
            this.labelEnd.TabIndex = 29;
            this.labelEnd.Text = "по";
            // 
            // dateTimePickerBegin
            // 
            this.dateTimePickerBegin.Enabled = false;
            this.dateTimePickerBegin.Location = new System.Drawing.Point(413, 231);
            this.dateTimePickerBegin.Name = "dateTimePickerBegin";
            this.dateTimePickerBegin.Size = new System.Drawing.Size(158, 20);
            this.dateTimePickerBegin.TabIndex = 26;
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Enabled = false;
            this.dateTimePickerEnd.Location = new System.Drawing.Point(413, 257);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(158, 20);
            this.dateTimePickerEnd.TabIndex = 27;
            // 
            // labelBegin
            // 
            this.labelBegin.AutoSize = true;
            this.labelBegin.Enabled = false;
            this.labelBegin.Location = new System.Drawing.Point(386, 234);
            this.labelBegin.Name = "labelBegin";
            this.labelBegin.Size = new System.Drawing.Size(13, 13);
            this.labelBegin.TabIndex = 28;
            this.labelBegin.Text = "c";
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(609, 354);
            this.Controls.Add(this.labelEnd);
            this.Controls.Add(this.dateTimePickerBegin);
            this.Controls.Add(this.dateTimePickerEnd);
            this.Controls.Add(this.labelBegin);
            this.Controls.Add(this.radioButtonForOIFailureByPeriod);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.buttonLoadListDevicesDifferentFiles);
            this.Controls.Add(this.textBoxFileNameListDevicesDifferentFiles);
            this.Controls.Add(this.radioButtonForListDevicesDifferentFiles);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonLoadFileListDevicesKoef);
            this.Controls.Add(this.textBoxFileNameDevicesKoef);
            this.Controls.Add(this.radioButtonForListDeviceKoef);
            this.Controls.Add(this.radioButtonForListNomenclature);
            this.Controls.Add(this.buttonLoadFileListNomenclature);
            this.Controls.Add(this.textBoxFileNameListNomenclature);
            this.Controls.Add(this.buttonLoadFileListDevices);
            this.Controls.Add(this.textBoxFileNameListDevice);
            this.Controls.Add(this.radioButtonForListDevice);
            this.Controls.Add(this.radioButtonForListPKI);
            this.Controls.Add(this.radioButtonForSelectedObject);
            this.Controls.Add(this.buttonLoadFileListPKI);
            this.Controls.Add(this.textBoxFileNameListPKI);
            this.Controls.Add(this.buttonMakeReport);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройка отчета";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.Button buttonLoadFileListPKI;
        private System.Windows.Forms.TextBox textBoxFileNameListPKI;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.RadioButton radioButtonForSelectedObject;
        private System.Windows.Forms.RadioButton radioButtonForListPKI;
        private System.Windows.Forms.RadioButton radioButtonForListDevice;
        private System.Windows.Forms.Button buttonLoadFileListDevices;
        private System.Windows.Forms.TextBox textBoxFileNameListDevice;
        private System.Windows.Forms.RadioButton radioButtonForListNomenclature;
        private System.Windows.Forms.Button buttonLoadFileListNomenclature;
        private System.Windows.Forms.TextBox textBoxFileNameListNomenclature;
        private System.Windows.Forms.Button buttonLoadFileListDevicesKoef;
        private System.Windows.Forms.TextBox textBoxFileNameDevicesKoef;
        private System.Windows.Forms.RadioButton radioButtonForListDeviceKoef;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.Button buttonLoadListDevicesDifferentFiles;
        private System.Windows.Forms.TextBox textBoxFileNameListDevicesDifferentFiles;
        private System.Windows.Forms.RadioButton radioButtonForListDevicesDifferentFiles;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton radioButtonForOIFailureByPeriod;
        private System.Windows.Forms.Label labelEnd;
        private System.Windows.Forms.DateTimePicker dateTimePickerBegin;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Label labelBegin;
    }
}