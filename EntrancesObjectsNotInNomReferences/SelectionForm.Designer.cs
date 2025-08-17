namespace EntrancesObjectsNotInNomReferences
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
            this.groupBoxNomenclatureReference = new System.Windows.Forms.GroupBox();
            this.radioButtonMaterials = new System.Windows.Forms.RadioButton();
            this.radioButtonOtherItems = new System.Windows.Forms.RadioButton();
            this.radioButtonStandartItems = new System.Windows.Forms.RadioButton();
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.comboBoxTypeReport = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxPathFolder = new System.Windows.Forms.TextBox();
            this.labelChangeFolder = new System.Windows.Forms.Label();
            this.buttonChangeFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBoxNomenclatureReference.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxNomenclatureReference
            // 
            this.groupBoxNomenclatureReference.Controls.Add(this.radioButtonMaterials);
            this.groupBoxNomenclatureReference.Controls.Add(this.radioButtonOtherItems);
            this.groupBoxNomenclatureReference.Controls.Add(this.radioButtonStandartItems);
            this.groupBoxNomenclatureReference.Location = new System.Drawing.Point(12, 80);
            this.groupBoxNomenclatureReference.Name = "groupBoxNomenclatureReference";
            this.groupBoxNomenclatureReference.Size = new System.Drawing.Size(546, 63);
            this.groupBoxNomenclatureReference.TabIndex = 0;
            this.groupBoxNomenclatureReference.TabStop = false;
            this.groupBoxNomenclatureReference.Text = "Номенклатурные справочники";
            // 
            // radioButtonMaterials
            // 
            this.radioButtonMaterials.AutoSize = true;
            this.radioButtonMaterials.Location = new System.Drawing.Point(407, 27);
            this.radioButtonMaterials.Name = "radioButtonMaterials";
            this.radioButtonMaterials.Size = new System.Drawing.Size(83, 17);
            this.radioButtonMaterials.TabIndex = 2;
            this.radioButtonMaterials.Text = "Материалы";
            this.radioButtonMaterials.UseVisualStyleBackColor = true;
            // 
            // radioButtonOtherItems
            // 
            this.radioButtonOtherItems.AutoSize = true;
            this.radioButtonOtherItems.Location = new System.Drawing.Point(210, 27);
            this.radioButtonOtherItems.Name = "radioButtonOtherItems";
            this.radioButtonOtherItems.Size = new System.Drawing.Size(107, 17);
            this.radioButtonOtherItems.TabIndex = 1;
            this.radioButtonOtherItems.Text = "Прочие изделия";
            this.radioButtonOtherItems.UseVisualStyleBackColor = true;
            // 
            // radioButtonStandartItems
            // 
            this.radioButtonStandartItems.AutoSize = true;
            this.radioButtonStandartItems.Checked = true;
            this.radioButtonStandartItems.Location = new System.Drawing.Point(17, 27);
            this.radioButtonStandartItems.Name = "radioButtonStandartItems";
            this.radioButtonStandartItems.Size = new System.Drawing.Size(137, 17);
            this.radioButtonStandartItems.TabIndex = 0;
            this.radioButtonStandartItems.TabStop = true;
            this.radioButtonStandartItems.Text = "Стандартные изделия";
            this.radioButtonStandartItems.UseVisualStyleBackColor = true;
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(12, 159);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(546, 23);
            this.buttonMakeReport.TabIndex = 1;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click_1);
            // 
            // comboBoxTypeReport
            // 
            this.comboBoxTypeReport.FormattingEnabled = true;
            this.comboBoxTypeReport.Items.AddRange(new object[] {
            "Отсутствующие объекты справочника",
            "Отсутствующие объекты справочника и их входимости",
            "Объекты справочника, соответствующие номенклатурным с примечанием",
            "Объекты справочника и их входимости, соответствующие номенклатурным с примечанием" +
                "",
            "Объекты справочника и их входимости, соответствующие номенклатурным с примечанием" +
                " (со списокм заказов)",
            "Заказы со списком прямых вхождений Стандартных изделий с примечаниями"});
            this.comboBoxTypeReport.Location = new System.Drawing.Point(80, 15);
            this.comboBoxTypeReport.Name = "comboBoxTypeReport";
            this.comboBoxTypeReport.Size = new System.Drawing.Size(478, 21);
            this.comboBoxTypeReport.TabIndex = 2;
            this.comboBoxTypeReport.SelectedIndexChanged += new System.EventHandler(this.comboBoxTypeReport_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Тип отчета";
            // 
            // textBoxPathFolder
            // 
            this.textBoxPathFolder.Location = new System.Drawing.Point(204, 47);
            this.textBoxPathFolder.Name = "textBoxPathFolder";
            this.textBoxPathFolder.Size = new System.Drawing.Size(312, 20);
            this.textBoxPathFolder.TabIndex = 4;
            this.textBoxPathFolder.Visible = false;
            this.textBoxPathFolder.TextChanged += new System.EventHandler(this.textBoxPathFolder_TextChanged);
            // 
            // labelChangeFolder
            // 
            this.labelChangeFolder.AutoSize = true;
            this.labelChangeFolder.Location = new System.Drawing.Point(12, 50);
            this.labelChangeFolder.Name = "labelChangeFolder";
            this.labelChangeFolder.Size = new System.Drawing.Size(186, 13);
            this.labelChangeFolder.TabIndex = 5;
            this.labelChangeFolder.Text = "Выбор папки для выгрузки файлов";
            this.labelChangeFolder.Visible = false;
            // 
            // buttonChangeFolder
            // 
            this.buttonChangeFolder.Location = new System.Drawing.Point(514, 45);
            this.buttonChangeFolder.Name = "buttonChangeFolder";
            this.buttonChangeFolder.Size = new System.Drawing.Size(44, 23);
            this.buttonChangeFolder.TabIndex = 6;
            this.buttonChangeFolder.Text = "...";
            this.buttonChangeFolder.UseVisualStyleBackColor = true;
            this.buttonChangeFolder.Visible = false;
            this.buttonChangeFolder.Click += new System.EventHandler(this.buttonChangeFolder_Click);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.HelpRequest += new System.EventHandler(this.folderBrowserDialog1_HelpRequest);
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(568, 199);
            this.Controls.Add(this.buttonChangeFolder);
            this.Controls.Add(this.labelChangeFolder);
            this.Controls.Add(this.textBoxPathFolder);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBoxTypeReport);
            this.Controls.Add(this.buttonMakeReport);
            this.Controls.Add(this.groupBoxNomenclatureReference);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Входимость объектов, отсутствующих в Номенклатурных справочниках";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBoxNomenclatureReference.ResumeLayout(false);
            this.groupBoxNomenclatureReference.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxNomenclatureReference;
        private System.Windows.Forms.RadioButton radioButtonMaterials;
        private System.Windows.Forms.RadioButton radioButtonOtherItems;
        private System.Windows.Forms.RadioButton radioButtonStandartItems;
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.ComboBox comboBoxTypeReport;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxPathFolder;
        private System.Windows.Forms.Label labelChangeFolder;
        private System.Windows.Forms.Button buttonChangeFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}