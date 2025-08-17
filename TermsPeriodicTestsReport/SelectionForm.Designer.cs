namespace TermsPeriodicTestsReport
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
            this.listViewFindObjects = new System.Windows.Forms.ListView();
            this.ID = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.NomObj = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.NomDesign = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.dateTimePickerOrder = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.SearchNameButton = new System.Windows.Forms.Button();
            this.textBoxNameOrDenotation = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // listViewFindObjects
            // 
            this.listViewFindObjects.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.ID,
            this.NomObj,
            this.NomDesign});
            this.listViewFindObjects.FullRowSelect = true;
            this.listViewFindObjects.GridLines = true;
            this.listViewFindObjects.Location = new System.Drawing.Point(13, 23);
            this.listViewFindObjects.Name = "listViewFindObjects";
            this.listViewFindObjects.Size = new System.Drawing.Size(637, 212);
            this.listViewFindObjects.TabIndex = 12;
            this.listViewFindObjects.UseCompatibleStateImageBehavior = false;
            this.listViewFindObjects.View = System.Windows.Forms.View.Details;
            this.listViewFindObjects.SelectedIndexChanged += new System.EventHandler(this.listViewFindObjects_SelectedIndexChanged);
            // 
            // ID
            // 
            this.ID.Text = "ID";
            this.ID.Width = 40;
            // 
            // NomObj
            // 
            this.NomObj.Text = "Наименование";
            this.NomObj.Width = 351;
            // 
            // NomDesign
            // 
            this.NomDesign.Text = "Обозначение";
            this.NomDesign.Width = 217;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listViewFindObjects);
            this.groupBox1.Location = new System.Drawing.Point(22, 114);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(665, 248);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Найденные объекты";
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Enabled = false;
            this.buttonMakeReport.Location = new System.Drawing.Point(22, 373);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(665, 23);
            this.buttonMakeReport.TabIndex = 14;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // dateTimePickerOrder
            // 
            this.dateTimePickerOrder.Location = new System.Drawing.Point(198, 62);
            this.dateTimePickerOrder.Name = "dateTimePickerOrder";
            this.dateTimePickerOrder.Size = new System.Drawing.Size(189, 20);
            this.dateTimePickerOrder.TabIndex = 15;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 13);
            this.label1.TabIndex = 16;
            this.label1.Text = "Дата сборки комплекса";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.textBoxNameOrDenotation);
            this.groupBox2.Controls.Add(this.SearchNameButton);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.dateTimePickerOrder);
            this.groupBox2.Location = new System.Drawing.Point(28, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(659, 93);
            this.groupBox2.TabIndex = 17;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Параметры для поиска";
            // 
            // SearchNameButton
            // 
            this.SearchNameButton.Location = new System.Drawing.Point(419, 33);
            this.SearchNameButton.Name = "SearchNameButton";
            this.SearchNameButton.Size = new System.Drawing.Size(225, 45);
            this.SearchNameButton.TabIndex = 17;
            this.SearchNameButton.Text = "Поиск";
            this.SearchNameButton.UseVisualStyleBackColor = true;
            this.SearchNameButton.Click += new System.EventHandler(this.SearchNameButton_Click);
            // 
            // textBoxNameOrDenotation
            // 
            this.textBoxNameOrDenotation.Location = new System.Drawing.Point(198, 33);
            this.textBoxNameOrDenotation.Name = "textBoxNameOrDenotation";
            this.textBoxNameOrDenotation.Size = new System.Drawing.Size(189, 20);
            this.textBoxNameOrDenotation.TabIndex = 18;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(175, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "Наименование или обозначение ";
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(705, 410);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.buttonMakeReport);
            this.Controls.Add(this.groupBox1);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет по срокам периодических испытаний";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ListView listViewFindObjects;
        private System.Windows.Forms.ColumnHeader ID;
        private System.Windows.Forms.ColumnHeader NomObj;
        private System.Windows.Forms.ColumnHeader NomDesign;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.DateTimePicker dateTimePickerOrder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxNameOrDenotation;
        private System.Windows.Forms.Button SearchNameButton;
    }
}