namespace Globus.PlanReportsExcel
{
    partial class ReportParametersForm
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
            this.OKButton = new System.Windows.Forms.Button();
            this.typeGroupBox = new System.Windows.Forms.GroupBox();
            this.listFormRadioButton = new System.Windows.Forms.RadioButton();
            this.tableFormRadioButton = new System.Windows.Forms.RadioButton();
            this.orientationGroupBox = new System.Windows.Forms.GroupBox();
            this.bookRadioButton = new System.Windows.Forms.RadioButton();
            this.albumRadioButton = new System.Windows.Forms.RadioButton();
            this.newPageGroupBox = new System.Windows.Forms.GroupBox();
            this.newPageRadioButton = new System.Windows.Forms.RadioButton();
            this.firstPageRadioButton = new System.Windows.Forms.RadioButton();
            this.typeGroupBox.SuspendLayout();
            this.orientationGroupBox.SuspendLayout();
            this.newPageGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // OKButton
            // 
            this.OKButton.Location = new System.Drawing.Point(147, 235);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(75, 23);
            this.OKButton.TabIndex = 0;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // typeGroupBox
            // 
            this.typeGroupBox.Controls.Add(this.tableFormRadioButton);
            this.typeGroupBox.Controls.Add(this.listFormRadioButton);
            this.typeGroupBox.Location = new System.Drawing.Point(12, 12);
            this.typeGroupBox.Name = "typeGroupBox";
            this.typeGroupBox.Size = new System.Drawing.Size(342, 76);
            this.typeGroupBox.TabIndex = 1;
            this.typeGroupBox.TabStop = false;
            this.typeGroupBox.Text = "Форма отчета";
            // 
            // listFormRadioButton
            // 
            this.listFormRadioButton.AutoSize = true;
            this.listFormRadioButton.Location = new System.Drawing.Point(6, 30);
            this.listFormRadioButton.Name = "listFormRadioButton";
            this.listFormRadioButton.Size = new System.Drawing.Size(299, 17);
            this.listFormRadioButton.TabIndex = 2;
            this.listFormRadioButton.Text = "Список работ (ДЗ отд.906, решение ОС, протокол ОС)";
            this.listFormRadioButton.UseVisualStyleBackColor = true;
            // 
            // tableFormRadioButton
            // 
            this.tableFormRadioButton.AutoSize = true;
            this.tableFormRadioButton.Checked = true;
            this.tableFormRadioButton.Location = new System.Drawing.Point(6, 53);
            this.tableFormRadioButton.Name = "tableFormRadioButton";
            this.tableFormRadioButton.Size = new System.Drawing.Size(241, 17);
            this.tableFormRadioButton.TabIndex = 3;
            this.tableFormRadioButton.TabStop = true;
            this.tableFormRadioButton.Text = "Табличная форма (графики, мероприятия)";
            this.tableFormRadioButton.UseVisualStyleBackColor = true;
            this.tableFormRadioButton.CheckedChanged += new System.EventHandler(this.tableFormRadioButton_CheckedChanged);
            // 
            // orientationGroupBox
            // 
            this.orientationGroupBox.Controls.Add(this.albumRadioButton);
            this.orientationGroupBox.Controls.Add(this.bookRadioButton);
            this.orientationGroupBox.Location = new System.Drawing.Point(12, 94);
            this.orientationGroupBox.Name = "orientationGroupBox";
            this.orientationGroupBox.Size = new System.Drawing.Size(342, 61);
            this.orientationGroupBox.TabIndex = 2;
            this.orientationGroupBox.TabStop = false;
            this.orientationGroupBox.Text = "Ориентация страницы";
            // 
            // bookRadioButton
            // 
            this.bookRadioButton.AutoSize = true;
            this.bookRadioButton.Location = new System.Drawing.Point(6, 29);
            this.bookRadioButton.Name = "bookRadioButton";
            this.bookRadioButton.Size = new System.Drawing.Size(154, 17);
            this.bookRadioButton.TabIndex = 0;
            this.bookRadioButton.Text = "А4 книжная (210х297 мм)";
            this.bookRadioButton.UseVisualStyleBackColor = true;
            // 
            // albumRadioButton
            // 
            this.albumRadioButton.AutoSize = true;
            this.albumRadioButton.Checked = true;
            this.albumRadioButton.Location = new System.Drawing.Point(166, 29);
            this.albumRadioButton.Name = "albumRadioButton";
            this.albumRadioButton.Size = new System.Drawing.Size(163, 17);
            this.albumRadioButton.TabIndex = 1;
            this.albumRadioButton.TabStop = true;
            this.albumRadioButton.Text = "А4 альбомная (297х210мм)";
            this.albumRadioButton.UseVisualStyleBackColor = true;
            this.albumRadioButton.CheckedChanged += new System.EventHandler(this.albumRadioButton_CheckedChanged);
            // 
            // newPageGroupBox
            // 
            this.newPageGroupBox.Controls.Add(this.firstPageRadioButton);
            this.newPageGroupBox.Controls.Add(this.newPageRadioButton);
            this.newPageGroupBox.Location = new System.Drawing.Point(12, 161);
            this.newPageGroupBox.Name = "newPageGroupBox";
            this.newPageGroupBox.Size = new System.Drawing.Size(342, 60);
            this.newPageGroupBox.TabIndex = 3;
            this.newPageGroupBox.TabStop = false;
            this.newPageGroupBox.Text = "Расположение списка (таблицы)";
            // 
            // newPageRadioButton
            // 
            this.newPageRadioButton.AutoSize = true;
            this.newPageRadioButton.Location = new System.Drawing.Point(6, 30);
            this.newPageRadioButton.Name = "newPageRadioButton";
            this.newPageRadioButton.Size = new System.Drawing.Size(122, 17);
            this.newPageRadioButton.TabIndex = 4;
            this.newPageRadioButton.TabStop = true;
            this.newPageRadioButton.Text = "На новой странице";
            this.newPageRadioButton.UseVisualStyleBackColor = true;
            // 
            // firstPageRadioButton
            // 
            this.firstPageRadioButton.AutoSize = true;
            this.firstPageRadioButton.Checked = true;
            this.firstPageRadioButton.Location = new System.Drawing.Point(166, 30);
            this.firstPageRadioButton.Name = "firstPageRadioButton";
            this.firstPageRadioButton.Size = new System.Drawing.Size(128, 17);
            this.firstPageRadioButton.TabIndex = 4;
            this.firstPageRadioButton.TabStop = true;
            this.firstPageRadioButton.Text = "На первой странице";
            this.firstPageRadioButton.UseVisualStyleBackColor = true;
            this.firstPageRadioButton.CheckedChanged += new System.EventHandler(this.firstPageRadioButton_CheckedChanged);
            // 
            // ReportParametersForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(366, 269);
            this.Controls.Add(this.newPageGroupBox);
            this.Controls.Add(this.orientationGroupBox);
            this.Controls.Add(this.typeGroupBox);
            this.Controls.Add(this.OKButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "ReportParametersForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Параметры отчета";
            this.Load += new System.EventHandler(this.ReportParametersForm_Load);
            this.typeGroupBox.ResumeLayout(false);
            this.typeGroupBox.PerformLayout();
            this.orientationGroupBox.ResumeLayout(false);
            this.orientationGroupBox.PerformLayout();
            this.newPageGroupBox.ResumeLayout(false);
            this.newPageGroupBox.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button OKButton;
        private System.Windows.Forms.GroupBox typeGroupBox;
        private System.Windows.Forms.RadioButton tableFormRadioButton;
        private System.Windows.Forms.RadioButton listFormRadioButton;
        private System.Windows.Forms.GroupBox orientationGroupBox;
        private System.Windows.Forms.RadioButton albumRadioButton;
        private System.Windows.Forms.RadioButton bookRadioButton;
        private System.Windows.Forms.GroupBox newPageGroupBox;
        private System.Windows.Forms.RadioButton firstPageRadioButton;
        private System.Windows.Forms.RadioButton newPageRadioButton;
    }
}