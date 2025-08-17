namespace SummaryReport
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
            this.btnMakeReport = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.fileNameTextBox = new System.Windows.Forms.TextBox();
            this.loadButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.listViewObjectsReport = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader10 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.buttonDelete = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBoxForOneObject = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chkAllTypes = new System.Windows.Forms.CheckBox();
            this.chkComponentProgram = new System.Windows.Forms.CheckBox();
            this.chkComplexProgram = new System.Windows.Forms.CheckBox();
            this.chkComplement = new System.Windows.Forms.CheckBox();
            this.chkMaterial = new System.Windows.Forms.CheckBox();
            this.chkOther = new System.Windows.Forms.CheckBox();
            this.chkStandart = new System.Windows.Forms.CheckBox();
            this.chkAssembly = new System.Windows.Forms.CheckBox();
            this.chkDetail = new System.Windows.Forms.CheckBox();
            this.chkComplex = new System.Windows.Forms.CheckBox();
            this.chkDocumentation = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnMakeReport
            // 
            this.btnMakeReport.Location = new System.Drawing.Point(12, 561);
            this.btnMakeReport.Name = "btnMakeReport";
            this.btnMakeReport.Size = new System.Drawing.Size(434, 23);
            this.btnMakeReport.TabIndex = 4;
            this.btnMakeReport.Text = "Сформировать отчет";
            this.btnMakeReport.UseVisualStyleBackColor = true;
            this.btnMakeReport.Click += new System.EventHandler(this.btnMakeReport_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(452, 561);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(86, 23);
            this.buttonClose.TabIndex = 5;
            this.buttonClose.Text = "Выход";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // fileNameTextBox
            // 
            this.fileNameTextBox.Location = new System.Drawing.Point(16, 24);
            this.fileNameTextBox.Name = "fileNameTextBox";
            this.fileNameTextBox.Size = new System.Drawing.Size(457, 20);
            this.fileNameTextBox.TabIndex = 6;
            // 
            // loadButton
            // 
            this.loadButton.Location = new System.Drawing.Point(467, 22);
            this.loadButton.Name = "loadButton";
            this.loadButton.Size = new System.Drawing.Size(40, 23);
            this.loadButton.TabIndex = 7;
            this.loadButton.Text = "...";
            this.loadButton.UseVisualStyleBackColor = true;
            this.loadButton.Click += new System.EventHandler(this.loadButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.loadButton);
            this.groupBox1.Controls.Add(this.fileNameTextBox);
            this.groupBox1.Location = new System.Drawing.Point(12, 38);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(526, 102);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Загрузка Excel-файла";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(221, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Обозначение, Наименование, Количество";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 58);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(341, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Столбцы таблицы файла должны содержать следующие данные: ";
            // 
            // listViewObjectsReport
            // 
            this.listViewObjectsReport.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader10});
            this.listViewObjectsReport.FullRowSelect = true;
            this.listViewObjectsReport.GridLines = true;
            this.listViewObjectsReport.Location = new System.Drawing.Point(7, 29);
            this.listViewObjectsReport.Name = "listViewObjectsReport";
            this.listViewObjectsReport.Size = new System.Drawing.Size(513, 167);
            this.listViewObjectsReport.TabIndex = 13;
            this.listViewObjectsReport.UseCompatibleStateImageBehavior = false;
            this.listViewObjectsReport.View = System.Windows.Forms.View.Details;
            this.listViewObjectsReport.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listViewObjectsReport_MouseDoubleClick);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "ID";
            this.columnHeader1.Width = 65;
            // 
            // columnHeader2
            // 
            this.columnHeader2.DisplayIndex = 2;
            this.columnHeader2.Text = "Наименование";
            this.columnHeader2.Width = 230;
            // 
            // columnHeader3
            // 
            this.columnHeader3.DisplayIndex = 1;
            this.columnHeader3.Text = "Обозначение";
            this.columnHeader3.Width = 155;
            // 
            // columnHeader10
            // 
            this.columnHeader10.Text = "Кол-во";
            this.columnHeader10.Width = 50;
            // 
            // buttonDelete
            // 
            this.buttonDelete.Location = new System.Drawing.Point(257, 207);
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.Size = new System.Drawing.Size(263, 23);
            this.buttonDelete.TabIndex = 14;
            this.buttonDelete.Text = "Исключить выделенный объект";
            this.buttonDelete.UseVisualStyleBackColor = true;
            this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.listViewObjectsReport);
            this.groupBox2.Controls.Add(this.buttonDelete);
            this.groupBox2.Location = new System.Drawing.Point(12, 146);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(526, 242);
            this.groupBox2.TabIndex = 15;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Объекты для отчета";
            // 
            // checkBoxForOneObject
            // 
            this.checkBoxForOneObject.AutoSize = true;
            this.checkBoxForOneObject.CheckAlign = System.Drawing.ContentAlignment.TopLeft;
            this.checkBoxForOneObject.Checked = true;
            this.checkBoxForOneObject.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxForOneObject.Location = new System.Drawing.Point(147, 15);
            this.checkBoxForOneObject.Name = "checkBoxForOneObject";
            this.checkBoxForOneObject.Size = new System.Drawing.Size(186, 17);
            this.checkBoxForOneObject.TabIndex = 17;
            this.checkBoxForOneObject.Text = "Ведомость для одного объекта";
            this.checkBoxForOneObject.UseVisualStyleBackColor = true;
            this.checkBoxForOneObject.CheckedChanged += new System.EventHandler(this.checkBoxForOneObject_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chkAllTypes);
            this.groupBox3.Controls.Add(this.chkComponentProgram);
            this.groupBox3.Controls.Add(this.chkComplexProgram);
            this.groupBox3.Controls.Add(this.chkComplement);
            this.groupBox3.Controls.Add(this.chkMaterial);
            this.groupBox3.Controls.Add(this.chkOther);
            this.groupBox3.Controls.Add(this.chkStandart);
            this.groupBox3.Controls.Add(this.chkAssembly);
            this.groupBox3.Controls.Add(this.chkDetail);
            this.groupBox3.Controls.Add(this.chkComplex);
            this.groupBox3.Controls.Add(this.chkDocumentation);
            this.groupBox3.Location = new System.Drawing.Point(12, 402);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(526, 140);
            this.groupBox3.TabIndex = 18;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Типы объектов";
            // 
            // chkAllTypes
            // 
            this.chkAllTypes.AutoSize = true;
            this.chkAllTypes.Checked = true;
            this.chkAllTypes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAllTypes.Location = new System.Drawing.Point(419, 19);
            this.chkAllTypes.Name = "chkAllTypes";
            this.chkAllTypes.Size = new System.Drawing.Size(73, 17);
            this.chkAllTypes.TabIndex = 10;
            this.chkAllTypes.Text = "Все типы";
            this.chkAllTypes.UseVisualStyleBackColor = true;
            this.chkAllTypes.CheckedChanged += new System.EventHandler(this.chkAllTypes_CheckedChanged);
            // 
            // chkComponentProgram
            // 
            this.chkComponentProgram.AutoSize = true;
            this.chkComponentProgram.Checked = true;
            this.chkComponentProgram.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkComponentProgram.Location = new System.Drawing.Point(205, 111);
            this.chkComponentProgram.Name = "chkComponentProgram";
            this.chkComponentProgram.Size = new System.Drawing.Size(158, 17);
            this.chkComponentProgram.TabIndex = 9;
            this.chkComponentProgram.Text = "Компоненты (программы)";
            this.chkComponentProgram.UseVisualStyleBackColor = true;
            // 
            // chkComplexProgram
            // 
            this.chkComplexProgram.AutoSize = true;
            this.chkComplexProgram.Checked = true;
            this.chkComplexProgram.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkComplexProgram.Location = new System.Drawing.Point(205, 88);
            this.chkComplexProgram.Name = "chkComplexProgram";
            this.chkComplexProgram.Size = new System.Drawing.Size(153, 17);
            this.chkComplexProgram.TabIndex = 8;
            this.chkComplexProgram.Text = "Комплексы (программы)";
            this.chkComplexProgram.UseVisualStyleBackColor = true;
            // 
            // chkComplement
            // 
            this.chkComplement.AutoSize = true;
            this.chkComplement.Checked = true;
            this.chkComplement.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkComplement.Location = new System.Drawing.Point(205, 65);
            this.chkComplement.Name = "chkComplement";
            this.chkComplement.Size = new System.Drawing.Size(84, 17);
            this.chkComplement.TabIndex = 7;
            this.chkComplement.Text = "Комплекты";
            this.chkComplement.UseVisualStyleBackColor = true;
            // 
            // chkMaterial
            // 
            this.chkMaterial.AutoSize = true;
            this.chkMaterial.Checked = true;
            this.chkMaterial.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMaterial.Location = new System.Drawing.Point(205, 42);
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
            this.chkOther.Location = new System.Drawing.Point(205, 19);
            this.chkOther.Name = "chkOther";
            this.chkOther.Size = new System.Drawing.Size(108, 17);
            this.chkOther.TabIndex = 5;
            this.chkOther.Text = "Прочие изделия";
            this.chkOther.UseVisualStyleBackColor = true;
            // 
            // chkStandart
            // 
            this.chkStandart.AutoSize = true;
            this.chkStandart.Checked = true;
            this.chkStandart.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkStandart.Location = new System.Drawing.Point(28, 111);
            this.chkStandart.Name = "chkStandart";
            this.chkStandart.Size = new System.Drawing.Size(138, 17);
            this.chkStandart.TabIndex = 4;
            this.chkStandart.Text = "Стандартные изделия";
            this.chkStandart.UseVisualStyleBackColor = true;
            // 
            // chkAssembly
            // 
            this.chkAssembly.AutoSize = true;
            this.chkAssembly.Checked = true;
            this.chkAssembly.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAssembly.Location = new System.Drawing.Point(28, 88);
            this.chkAssembly.Name = "chkAssembly";
            this.chkAssembly.Size = new System.Drawing.Size(129, 17);
            this.chkAssembly.TabIndex = 3;
            this.chkAssembly.Text = "Сборочные единицы";
            this.chkAssembly.UseVisualStyleBackColor = true;
            // 
            // chkDetail
            // 
            this.chkDetail.AutoSize = true;
            this.chkDetail.Checked = true;
            this.chkDetail.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDetail.Location = new System.Drawing.Point(28, 65);
            this.chkDetail.Name = "chkDetail";
            this.chkDetail.Size = new System.Drawing.Size(64, 17);
            this.chkDetail.TabIndex = 2;
            this.chkDetail.Text = "Детали";
            this.chkDetail.UseVisualStyleBackColor = true;
            // 
            // chkComplex
            // 
            this.chkComplex.AutoSize = true;
            this.chkComplex.Checked = true;
            this.chkComplex.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkComplex.Location = new System.Drawing.Point(28, 42);
            this.chkComplex.Name = "chkComplex";
            this.chkComplex.Size = new System.Drawing.Size(85, 17);
            this.chkComplex.TabIndex = 1;
            this.chkComplex.Text = "Комплексы";
            this.chkComplex.UseVisualStyleBackColor = true;
            // 
            // chkDocumentation
            // 
            this.chkDocumentation.AutoSize = true;
            this.chkDocumentation.Checked = true;
            this.chkDocumentation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDocumentation.Location = new System.Drawing.Point(28, 19);
            this.chkDocumentation.Name = "chkDocumentation";
            this.chkDocumentation.Size = new System.Drawing.Size(101, 17);
            this.chkDocumentation.TabIndex = 0;
            this.chkDocumentation.Text = "Документация";
            this.chkDocumentation.UseVisualStyleBackColor = true;
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(550, 597);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.checkBoxForOneObject);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.btnMakeReport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет - Сводная ведомость";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnMakeReport;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.TextBox fileNameTextBox;
        private System.Windows.Forms.Button loadButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListView listViewObjectsReport;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader10;
        private System.Windows.Forms.Button buttonDelete;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBoxForOneObject;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox chkAllTypes;
        private System.Windows.Forms.CheckBox chkComponentProgram;
        private System.Windows.Forms.CheckBox chkComplexProgram;
        private System.Windows.Forms.CheckBox chkComplement;
        private System.Windows.Forms.CheckBox chkMaterial;
        private System.Windows.Forms.CheckBox chkOther;
        private System.Windows.Forms.CheckBox chkStandart;
        private System.Windows.Forms.CheckBox chkAssembly;
        private System.Windows.Forms.CheckBox chkDetail;
        private System.Windows.Forms.CheckBox chkComplex;
        private System.Windows.Forms.CheckBox chkDocumentation;
    }
}