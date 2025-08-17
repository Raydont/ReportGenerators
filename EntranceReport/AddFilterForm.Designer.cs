namespace EntranceReport
{
    partial class AddFilterForm
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
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.addFilterButton = new System.Windows.Forms.Button();
            this.nameGroupBox = new System.Windows.Forms.GroupBox();
            this.textBoxAddLikeName = new System.Windows.Forms.TextBox();
            this.AddConitionsNameCheckBox = new System.Windows.Forms.CheckBox();
            this.denotationGroupBox = new System.Windows.Forms.GroupBox();
            this.textBoxAddLikeDenotation = new System.Windows.Forms.TextBox();
            this.AddConitionsDenotationCheckBox = new System.Windows.Forms.CheckBox();
            this.filterListView = new System.Windows.Forms.ListView();
            this.columnHeaderNumber = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderDenotation = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.buttonClose = new System.Windows.Forms.Button();
            this.btnMakeReport = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBoxDenotation = new System.Windows.Forms.CheckBox();
            this.checkBoxName = new System.Windows.Forms.CheckBox();
            this.tabControlEntranceReport = new System.Windows.Forms.TabControl();
            this.tabPageChoiceTypesAndColumns = new System.Windows.Forms.TabPage();
            this.columnSelectionGroupBox = new System.Windows.Forms.GroupBox();
            this.chkUnitMeasure = new System.Windows.Forms.CheckBox();
            this.chkCountOnProduct = new System.Windows.Forms.CheckBox();
            this.chkCountOnKnot = new System.Windows.Forms.CheckBox();
            this.chkAllObjects = new System.Windows.Forms.CheckBox();
            this.chkRemarks = new System.Windows.Forms.CheckBox();
            this.chkCount = new System.Windows.Forms.CheckBox();
            this.chkName = new System.Windows.Forms.CheckBox();
            this.chkDenotation = new System.Windows.Forms.CheckBox();
            this.chkFormat = new System.Windows.Forms.CheckBox();
            this.chkZone = new System.Windows.Forms.CheckBox();
            this.chkPosition = new System.Windows.Forms.CheckBox();
            this.tabPageSearchConditions = new System.Windows.Forms.TabPage();
            this.groupBox3.SuspendLayout();
            this.nameGroupBox.SuspendLayout();
            this.denotationGroupBox.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabControlEntranceReport.SuspendLayout();
            this.tabPageChoiceTypesAndColumns.SuspendLayout();
            this.columnSelectionGroupBox.SuspendLayout();
            this.tabPageSearchConditions.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.addFilterButton);
            this.groupBox3.Controls.Add(this.nameGroupBox);
            this.groupBox3.Controls.Add(this.AddConitionsNameCheckBox);
            this.groupBox3.Controls.Add(this.denotationGroupBox);
            this.groupBox3.Controls.Add(this.AddConitionsDenotationCheckBox);
            this.groupBox3.Location = new System.Drawing.Point(8, 144);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(758, 138);
            this.groupBox3.TabIndex = 25;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Добавить условие";
            // 
            // addFilterButton
            // 
            this.addFilterButton.Enabled = false;
            this.addFilterButton.Location = new System.Drawing.Point(585, 105);
            this.addFilterButton.Name = "addFilterButton";
            this.addFilterButton.Size = new System.Drawing.Size(154, 23);
            this.addFilterButton.TabIndex = 4;
            this.addFilterButton.Text = "Добавить";
            this.addFilterButton.UseVisualStyleBackColor = true;
            this.addFilterButton.Click += new System.EventHandler(this.addFilterButton_Click);
            // 
            // nameGroupBox
            // 
            this.nameGroupBox.Controls.Add(this.textBoxAddLikeName);
            this.nameGroupBox.Enabled = false;
            this.nameGroupBox.Location = new System.Drawing.Point(21, 44);
            this.nameGroupBox.Name = "nameGroupBox";
            this.nameGroupBox.Size = new System.Drawing.Size(346, 55);
            this.nameGroupBox.TabIndex = 0;
            this.nameGroupBox.TabStop = false;
            this.nameGroupBox.Text = "Добавление условия содержит Наименование";
            // 
            // textBoxAddLikeName
            // 
            this.textBoxAddLikeName.Location = new System.Drawing.Point(40, 24);
            this.textBoxAddLikeName.Name = "textBoxAddLikeName";
            this.textBoxAddLikeName.Size = new System.Drawing.Size(283, 20);
            this.textBoxAddLikeName.TabIndex = 0;
            // 
            // AddConitionsNameCheckBox
            // 
            this.AddConitionsNameCheckBox.AutoSize = true;
            this.AddConitionsNameCheckBox.Location = new System.Drawing.Point(21, 19);
            this.AddConitionsNameCheckBox.Name = "AddConitionsNameCheckBox";
            this.AddConitionsNameCheckBox.Size = new System.Drawing.Size(265, 17);
            this.AddConitionsNameCheckBox.TabIndex = 2;
            this.AddConitionsNameCheckBox.Text = "Добавить условие \"Содержит Наименование\" ";
            this.AddConitionsNameCheckBox.UseVisualStyleBackColor = true;
            this.AddConitionsNameCheckBox.CheckedChanged += new System.EventHandler(this.AddConitionsNameCheckBox_CheckedChanged);
            // 
            // denotationGroupBox
            // 
            this.denotationGroupBox.Controls.Add(this.textBoxAddLikeDenotation);
            this.denotationGroupBox.Enabled = false;
            this.denotationGroupBox.Location = new System.Drawing.Point(403, 44);
            this.denotationGroupBox.Name = "denotationGroupBox";
            this.denotationGroupBox.Size = new System.Drawing.Size(336, 55);
            this.denotationGroupBox.TabIndex = 1;
            this.denotationGroupBox.TabStop = false;
            this.denotationGroupBox.Text = "Добавление условия содержит Обозначение";
            // 
            // textBoxAddLikeDenotation
            // 
            this.textBoxAddLikeDenotation.Location = new System.Drawing.Point(36, 24);
            this.textBoxAddLikeDenotation.Name = "textBoxAddLikeDenotation";
            this.textBoxAddLikeDenotation.Size = new System.Drawing.Size(283, 20);
            this.textBoxAddLikeDenotation.TabIndex = 1;
            // 
            // AddConitionsDenotationCheckBox
            // 
            this.AddConitionsDenotationCheckBox.AutoSize = true;
            this.AddConitionsDenotationCheckBox.Location = new System.Drawing.Point(403, 19);
            this.AddConitionsDenotationCheckBox.Name = "AddConitionsDenotationCheckBox";
            this.AddConitionsDenotationCheckBox.Size = new System.Drawing.Size(256, 17);
            this.AddConitionsDenotationCheckBox.TabIndex = 3;
            this.AddConitionsDenotationCheckBox.Text = "Добавить условие \"Содержит Обозначение\" ";
            this.AddConitionsDenotationCheckBox.UseVisualStyleBackColor = true;
            this.AddConitionsDenotationCheckBox.CheckedChanged += new System.EventHandler(this.AddConitionsDenotationCheckBox_CheckedChanged);
            // 
            // filterListView
            // 
            this.filterListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderNumber,
            this.columnHeaderName,
            this.columnHeaderDenotation});
            this.filterListView.FullRowSelect = true;
            this.filterListView.GridLines = true;
            this.filterListView.Location = new System.Drawing.Point(42, 15);
            this.filterListView.Name = "filterListView";
            this.filterListView.Size = new System.Drawing.Size(688, 113);
            this.filterListView.TabIndex = 24;
            this.filterListView.UseCompatibleStateImageBehavior = false;
            this.filterListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeaderNumber
            // 
            this.columnHeaderNumber.Text = "№";
            // 
            // columnHeaderName
            // 
            this.columnHeaderName.Text = "Наименование";
            this.columnHeaderName.Width = 300;
            // 
            // columnHeaderDenotation
            // 
            this.columnHeaderDenotation.Text = "Обозначение";
            this.columnHeaderDenotation.Width = 300;
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(689, 340);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(111, 23);
            this.buttonClose.TabIndex = 27;
            this.buttonClose.Text = "Выход";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // btnMakeReport
            // 
            this.btnMakeReport.Location = new System.Drawing.Point(505, 340);
            this.btnMakeReport.Name = "btnMakeReport";
            this.btnMakeReport.Size = new System.Drawing.Size(179, 23);
            this.btnMakeReport.TabIndex = 26;
            this.btnMakeReport.Text = "Сформировать отчет";
            this.btnMakeReport.UseVisualStyleBackColor = true;
            this.btnMakeReport.Click += new System.EventHandler(this.btnMakeReport_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkAllTypes);
            this.groupBox2.Controls.Add(this.chkComponentProgram);
            this.groupBox2.Controls.Add(this.chkComplexProgram);
            this.groupBox2.Controls.Add(this.chkComplement);
            this.groupBox2.Controls.Add(this.chkMaterial);
            this.groupBox2.Controls.Add(this.chkOther);
            this.groupBox2.Controls.Add(this.chkStandart);
            this.groupBox2.Controls.Add(this.chkAssembly);
            this.groupBox2.Controls.Add(this.chkDetail);
            this.groupBox2.Controls.Add(this.chkComplex);
            this.groupBox2.Controls.Add(this.chkDocumentation);
            this.groupBox2.Location = new System.Drawing.Point(9, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(754, 130);
            this.groupBox2.TabIndex = 28;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Типы объектов";
            // 
            // chkAllTypes
            // 
            this.chkAllTypes.AutoSize = true;
            this.chkAllTypes.Checked = true;
            this.chkAllTypes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAllTypes.Location = new System.Drawing.Point(637, 25);
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
            this.chkComponentProgram.Location = new System.Drawing.Point(390, 48);
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
            this.chkComplexProgram.Location = new System.Drawing.Point(390, 25);
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
            this.chkComplement.Location = new System.Drawing.Point(192, 94);
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
            this.chkMaterial.Location = new System.Drawing.Point(192, 71);
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
            this.chkOther.Location = new System.Drawing.Point(192, 48);
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
            this.chkStandart.Location = new System.Drawing.Point(192, 25);
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
            this.chkAssembly.Location = new System.Drawing.Point(34, 94);
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
            this.chkDetail.Location = new System.Drawing.Point(34, 71);
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
            this.chkComplex.Location = new System.Drawing.Point(34, 48);
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
            this.chkDocumentation.Location = new System.Drawing.Point(34, 25);
            this.chkDocumentation.Name = "chkDocumentation";
            this.chkDocumentation.Size = new System.Drawing.Size(101, 17);
            this.chkDocumentation.TabIndex = 0;
            this.chkDocumentation.Text = "Документация";
            this.chkDocumentation.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBoxDenotation);
            this.groupBox1.Controls.Add(this.checkBoxName);
            this.groupBox1.Location = new System.Drawing.Point(457, 19);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(149, 68);
            this.groupBox1.TabIndex = 29;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Входимость";
            // 
            // checkBoxDenotation
            // 
            this.checkBoxDenotation.AutoSize = true;
            this.checkBoxDenotation.Checked = true;
            this.checkBoxDenotation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxDenotation.Location = new System.Drawing.Point(29, 42);
            this.checkBoxDenotation.Name = "checkBoxDenotation";
            this.checkBoxDenotation.Size = new System.Drawing.Size(93, 17);
            this.checkBoxDenotation.TabIndex = 1;
            this.checkBoxDenotation.Text = "Обозначение";
            this.checkBoxDenotation.UseVisualStyleBackColor = true;
            // 
            // checkBoxName
            // 
            this.checkBoxName.AutoSize = true;
            this.checkBoxName.Location = new System.Drawing.Point(29, 19);
            this.checkBoxName.Name = "checkBoxName";
            this.checkBoxName.Size = new System.Drawing.Size(102, 17);
            this.checkBoxName.TabIndex = 0;
            this.checkBoxName.Text = "Наименование";
            this.checkBoxName.UseVisualStyleBackColor = true;
            // 
            // tabControlEntranceReport
            // 
            this.tabControlEntranceReport.Controls.Add(this.tabPageChoiceTypesAndColumns);
            this.tabControlEntranceReport.Controls.Add(this.tabPageSearchConditions);
            this.tabControlEntranceReport.Location = new System.Drawing.Point(20, 12);
            this.tabControlEntranceReport.Name = "tabControlEntranceReport";
            this.tabControlEntranceReport.SelectedIndex = 0;
            this.tabControlEntranceReport.Size = new System.Drawing.Size(784, 322);
            this.tabControlEntranceReport.TabIndex = 30;
            // 
            // tabPageChoiceTypesAndColumns
            // 
            this.tabPageChoiceTypesAndColumns.Controls.Add(this.columnSelectionGroupBox);
            this.tabPageChoiceTypesAndColumns.Controls.Add(this.groupBox2);
            this.tabPageChoiceTypesAndColumns.Location = new System.Drawing.Point(4, 22);
            this.tabPageChoiceTypesAndColumns.Name = "tabPageChoiceTypesAndColumns";
            this.tabPageChoiceTypesAndColumns.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageChoiceTypesAndColumns.Size = new System.Drawing.Size(776, 296);
            this.tabPageChoiceTypesAndColumns.TabIndex = 0;
            this.tabPageChoiceTypesAndColumns.Text = "Выбор типов и колонок для отчета";
            this.tabPageChoiceTypesAndColumns.UseVisualStyleBackColor = true;
            // 
            // columnSelectionGroupBox
            // 
            this.columnSelectionGroupBox.Controls.Add(this.chkUnitMeasure);
            this.columnSelectionGroupBox.Controls.Add(this.chkCountOnProduct);
            this.columnSelectionGroupBox.Controls.Add(this.chkCountOnKnot);
            this.columnSelectionGroupBox.Controls.Add(this.chkAllObjects);
            this.columnSelectionGroupBox.Controls.Add(this.groupBox1);
            this.columnSelectionGroupBox.Controls.Add(this.chkRemarks);
            this.columnSelectionGroupBox.Controls.Add(this.chkCount);
            this.columnSelectionGroupBox.Controls.Add(this.chkName);
            this.columnSelectionGroupBox.Controls.Add(this.chkDenotation);
            this.columnSelectionGroupBox.Controls.Add(this.chkFormat);
            this.columnSelectionGroupBox.Controls.Add(this.chkZone);
            this.columnSelectionGroupBox.Controls.Add(this.chkPosition);
            this.columnSelectionGroupBox.Location = new System.Drawing.Point(9, 163);
            this.columnSelectionGroupBox.Name = "columnSelectionGroupBox";
            this.columnSelectionGroupBox.Size = new System.Drawing.Size(755, 116);
            this.columnSelectionGroupBox.TabIndex = 30;
            this.columnSelectionGroupBox.TabStop = false;
            this.columnSelectionGroupBox.Text = "Поля спецификации";
            // 
            // chkUnitMeasure
            // 
            this.chkUnitMeasure.AutoSize = true;
            this.chkUnitMeasure.Location = new System.Drawing.Point(143, 70);
            this.chkUnitMeasure.Name = "chkUnitMeasure";
            this.chkUnitMeasure.Size = new System.Drawing.Size(128, 17);
            this.chkUnitMeasure.TabIndex = 30;
            this.chkUnitMeasure.Text = "Единица измерения";
            this.chkUnitMeasure.UseVisualStyleBackColor = true;
            // 
            // chkCountOnProduct
            // 
            this.chkCountOnProduct.AutoSize = true;
            this.chkCountOnProduct.Checked = true;
            this.chkCountOnProduct.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCountOnProduct.Location = new System.Drawing.Point(297, 47);
            this.chkCountOnProduct.Name = "chkCountOnProduct";
            this.chkCountOnProduct.Size = new System.Drawing.Size(145, 17);
            this.chkCountOnProduct.TabIndex = 15;
            this.chkCountOnProduct.Text = "Количество на изделие";
            this.chkCountOnProduct.UseVisualStyleBackColor = true;
            // 
            // chkCountOnKnot
            // 
            this.chkCountOnKnot.AutoSize = true;
            this.chkCountOnKnot.Checked = true;
            this.chkCountOnKnot.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCountOnKnot.Location = new System.Drawing.Point(297, 24);
            this.chkCountOnKnot.Name = "chkCountOnKnot";
            this.chkCountOnKnot.Size = new System.Drawing.Size(126, 17);
            this.chkCountOnKnot.TabIndex = 14;
            this.chkCountOnKnot.Text = "Количество на узел";
            this.chkCountOnKnot.UseVisualStyleBackColor = true;
            // 
            // chkAllObjects
            // 
            this.chkAllObjects.AutoSize = true;
            this.chkAllObjects.Location = new System.Drawing.Point(637, 24);
            this.chkAllObjects.Name = "chkAllObjects";
            this.chkAllObjects.Size = new System.Drawing.Size(92, 17);
            this.chkAllObjects.TabIndex = 8;
            this.chkAllObjects.Text = "Все объекты";
            this.chkAllObjects.UseVisualStyleBackColor = true;
            this.chkAllObjects.CheckedChanged += new System.EventHandler(this.chkAllObjects_CheckedChanged);
            // 
            // chkRemarks
            // 
            this.chkRemarks.AutoSize = true;
            this.chkRemarks.Location = new System.Drawing.Point(143, 47);
            this.chkRemarks.Name = "chkRemarks";
            this.chkRemarks.Size = new System.Drawing.Size(89, 17);
            this.chkRemarks.TabIndex = 6;
            this.chkRemarks.Text = "Примечание";
            this.chkRemarks.UseVisualStyleBackColor = true;
            // 
            // chkCount
            // 
            this.chkCount.AutoSize = true;
            this.chkCount.Location = new System.Drawing.Point(143, 93);
            this.chkCount.Name = "chkCount";
            this.chkCount.Size = new System.Drawing.Size(85, 17);
            this.chkCount.TabIndex = 5;
            this.chkCount.Text = "Количество";
            this.chkCount.UseVisualStyleBackColor = true;
            this.chkCount.Visible = false;
            // 
            // chkName
            // 
            this.chkName.AutoSize = true;
            this.chkName.Checked = true;
            this.chkName.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkName.Location = new System.Drawing.Point(143, 24);
            this.chkName.Name = "chkName";
            this.chkName.Size = new System.Drawing.Size(102, 17);
            this.chkName.TabIndex = 4;
            this.chkName.Text = "Наименование";
            this.chkName.UseVisualStyleBackColor = true;
            // 
            // chkDenotation
            // 
            this.chkDenotation.AutoSize = true;
            this.chkDenotation.Checked = true;
            this.chkDenotation.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDenotation.Location = new System.Drawing.Point(23, 93);
            this.chkDenotation.Name = "chkDenotation";
            this.chkDenotation.Size = new System.Drawing.Size(93, 17);
            this.chkDenotation.TabIndex = 3;
            this.chkDenotation.Text = "Обозначение";
            this.chkDenotation.UseVisualStyleBackColor = true;
            // 
            // chkFormat
            // 
            this.chkFormat.AutoSize = true;
            this.chkFormat.Checked = true;
            this.chkFormat.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFormat.Location = new System.Drawing.Point(23, 70);
            this.chkFormat.Name = "chkFormat";
            this.chkFormat.Size = new System.Drawing.Size(68, 17);
            this.chkFormat.TabIndex = 2;
            this.chkFormat.Text = "Формат";
            this.chkFormat.UseVisualStyleBackColor = true;
            // 
            // chkZone
            // 
            this.chkZone.AutoSize = true;
            this.chkZone.Location = new System.Drawing.Point(23, 47);
            this.chkZone.Name = "chkZone";
            this.chkZone.Size = new System.Drawing.Size(51, 17);
            this.chkZone.TabIndex = 1;
            this.chkZone.Text = "Зона";
            this.chkZone.UseVisualStyleBackColor = true;
            // 
            // chkPosition
            // 
            this.chkPosition.AutoSize = true;
            this.chkPosition.Location = new System.Drawing.Point(23, 24);
            this.chkPosition.Name = "chkPosition";
            this.chkPosition.Size = new System.Drawing.Size(70, 17);
            this.chkPosition.TabIndex = 0;
            this.chkPosition.Text = "Позиция";
            this.chkPosition.UseVisualStyleBackColor = true;
            // 
            // tabPageSearchConditions
            // 
            this.tabPageSearchConditions.Controls.Add(this.filterListView);
            this.tabPageSearchConditions.Controls.Add(this.groupBox3);
            this.tabPageSearchConditions.Location = new System.Drawing.Point(4, 22);
            this.tabPageSearchConditions.Name = "tabPageSearchConditions";
            this.tabPageSearchConditions.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageSearchConditions.Size = new System.Drawing.Size(776, 296);
            this.tabPageSearchConditions.TabIndex = 1;
            this.tabPageSearchConditions.Text = "Условия поиска";
            this.tabPageSearchConditions.UseVisualStyleBackColor = true;
            // 
            // AddFilterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(811, 377);
            this.Controls.Add(this.tabControlEntranceReport);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.btnMakeReport);
            this.Name = "AddFilterForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ведомость входимостей";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.AddFilterForm_FormClosed);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.nameGroupBox.ResumeLayout(false);
            this.nameGroupBox.PerformLayout();
            this.denotationGroupBox.ResumeLayout(false);
            this.denotationGroupBox.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControlEntranceReport.ResumeLayout(false);
            this.tabPageChoiceTypesAndColumns.ResumeLayout(false);
            this.columnSelectionGroupBox.ResumeLayout(false);
            this.columnSelectionGroupBox.PerformLayout();
            this.tabPageSearchConditions.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button addFilterButton;
        private System.Windows.Forms.GroupBox nameGroupBox;
        private System.Windows.Forms.TextBox textBoxAddLikeName;
        private System.Windows.Forms.CheckBox AddConitionsNameCheckBox;
        private System.Windows.Forms.GroupBox denotationGroupBox;
        private System.Windows.Forms.TextBox textBoxAddLikeDenotation;
        private System.Windows.Forms.CheckBox AddConitionsDenotationCheckBox;
        private System.Windows.Forms.ListView filterListView;
        private System.Windows.Forms.ColumnHeader columnHeaderNumber;
        private System.Windows.Forms.ColumnHeader columnHeaderName;
        private System.Windows.Forms.ColumnHeader columnHeaderDenotation;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Button btnMakeReport;
        private System.Windows.Forms.GroupBox groupBox2;
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
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBoxDenotation;
        private System.Windows.Forms.CheckBox checkBoxName;
        private System.Windows.Forms.TabControl tabControlEntranceReport;
        private System.Windows.Forms.TabPage tabPageChoiceTypesAndColumns;
        private System.Windows.Forms.TabPage tabPageSearchConditions;
        private System.Windows.Forms.GroupBox columnSelectionGroupBox;
        private System.Windows.Forms.CheckBox chkCountOnProduct;
        private System.Windows.Forms.CheckBox chkCountOnKnot;
        private System.Windows.Forms.CheckBox chkAllObjects;
        private System.Windows.Forms.CheckBox chkRemarks;
        private System.Windows.Forms.CheckBox chkName;
        private System.Windows.Forms.CheckBox chkDenotation;
        private System.Windows.Forms.CheckBox chkFormat;
        private System.Windows.Forms.CheckBox chkZone;
        private System.Windows.Forms.CheckBox chkPosition;
        private System.Windows.Forms.CheckBox chkUnitMeasure;
        private System.Windows.Forms.CheckBox chkCount;
    }
}