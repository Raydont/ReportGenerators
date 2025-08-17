namespace TransitTimeProcess
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.radioButtonForSelectedObject = new System.Windows.Forms.RadioButton();
            this.radioButtonForList = new System.Windows.Forms.RadioButton();
            this.dateTimePickerBegin = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(29, 136);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(455, 23);
            this.buttonMakeReport.TabIndex = 0;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // radioButtonForSelectedObject
            // 
            this.radioButtonForSelectedObject.AutoSize = true;
            this.radioButtonForSelectedObject.Checked = true;
            this.radioButtonForSelectedObject.Location = new System.Drawing.Point(39, 27);
            this.radioButtonForSelectedObject.Name = "radioButtonForSelectedObject";
            this.radioButtonForSelectedObject.Size = new System.Drawing.Size(155, 17);
            this.radioButtonForSelectedObject.TabIndex = 8;
            this.radioButtonForSelectedObject.TabStop = true;
            this.radioButtonForSelectedObject.Text = "Для выбранного объекта";
            this.radioButtonForSelectedObject.UseVisualStyleBackColor = true;
            this.radioButtonForSelectedObject.CheckedChanged += new System.EventHandler(this.radioButtonForSelectedObject_CheckedChanged);
            // 
            // radioButtonForList
            // 
            this.radioButtonForList.AutoSize = true;
            this.radioButtonForList.Location = new System.Drawing.Point(39, 58);
            this.radioButtonForList.Name = "radioButtonForList";
            this.radioButtonForList.Size = new System.Drawing.Size(261, 17);
            this.radioButtonForList.TabIndex = 9;
            this.radioButtonForList.Text = "Для списка объектов по выбранному периоду";
            this.radioButtonForList.UseVisualStyleBackColor = true;
            this.radioButtonForList.CheckedChanged += new System.EventHandler(this.radioButtonForListPKI_CheckedChanged);
            // 
            // dateTimePickerBegin
            // 
            this.dateTimePickerBegin.Enabled = false;
            this.dateTimePickerBegin.Location = new System.Drawing.Point(340, 58);
            this.dateTimePickerBegin.Name = "dateTimePickerBegin";
            this.dateTimePickerBegin.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerBegin.TabIndex = 10;
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Enabled = false;
            this.dateTimePickerEnd.Location = new System.Drawing.Point(340, 94);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(144, 20);
            this.dateTimePickerEnd.TabIndex = 11;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(309, 61);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(13, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "c";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(308, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(19, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "по";
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(520, 179);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePickerEnd);
            this.Controls.Add(this.dateTimePickerBegin);
            this.Controls.Add(this.radioButtonForList);
            this.Controls.Add(this.radioButtonForSelectedObject);
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
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.RadioButton radioButtonForSelectedObject;
        private System.Windows.Forms.RadioButton radioButtonForList;
        private System.Windows.Forms.DateTimePicker dateTimePickerBegin;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}