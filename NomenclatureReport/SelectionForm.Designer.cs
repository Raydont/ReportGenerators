namespace NomenclatureReport
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
            this.chkCodeless = new System.Windows.Forms.CheckBox();
            this.chkCode = new System.Windows.Forms.CheckBox();
            this.chkAllTypes = new System.Windows.Forms.CheckBox();
            this.chkComplement = new System.Windows.Forms.CheckBox();
            this.chkMaterial = new System.Windows.Forms.CheckBox();
            this.chkOther = new System.Windows.Forms.CheckBox();
            this.chkStandart = new System.Windows.Forms.CheckBox();
            this.chkAssembly = new System.Windows.Forms.CheckBox();
            this.chkDetail = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnMakeReport = new System.Windows.Forms.Button();
            this.buttonClose = new System.Windows.Forms.Button();
            this.dgvProducts = new System.Windows.Forms.DataGridView();
            this.dgvAddedProducts = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.progressLabel = new System.Windows.Forms.Label();
            this.deleteComplexButton = new System.Windows.Forms.Button();
            this.clearDBButton = new System.Windows.Forms.Button();
            this.addProductButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBoxComplement = new System.Windows.Forms.CheckBox();
            this.checkBoxComplex = new System.Windows.Forms.CheckBox();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProducts)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAddedProducts)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.chkCodeless);
            this.groupBox2.Controls.Add(this.chkCode);
            this.groupBox2.Controls.Add(this.chkAllTypes);
            this.groupBox2.Controls.Add(this.chkComplement);
            this.groupBox2.Controls.Add(this.chkMaterial);
            this.groupBox2.Controls.Add(this.chkOther);
            this.groupBox2.Controls.Add(this.chkStandart);
            this.groupBox2.Controls.Add(this.chkAssembly);
            this.groupBox2.Controls.Add(this.chkDetail);
            this.groupBox2.Location = new System.Drawing.Point(764, 332);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(182, 246);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Типы объектов";
            // 
            // chkCodeless
            // 
            this.chkCodeless.AutoSize = true;
            this.chkCodeless.Enabled = false;
            this.chkCodeless.Location = new System.Drawing.Point(37, 115);
            this.chkCodeless.Name = "chkCodeless";
            this.chkCodeless.Size = new System.Drawing.Size(112, 17);
            this.chkCodeless.TabIndex = 12;
            this.chkCodeless.Text = "Нешифрованные";
            this.chkCodeless.UseVisualStyleBackColor = true;
            this.chkCodeless.CheckedChanged += new System.EventHandler(this.chkCodeless_CheckedChanged);
            // 
            // chkCode
            // 
            this.chkCode.AutoSize = true;
            this.chkCode.Enabled = false;
            this.chkCode.Location = new System.Drawing.Point(37, 92);
            this.chkCode.Name = "chkCode";
            this.chkCode.Size = new System.Drawing.Size(99, 17);
            this.chkCode.TabIndex = 11;
            this.chkCode.Text = "Шифрованные";
            this.chkCode.UseVisualStyleBackColor = true;
            this.chkCode.CheckedChanged += new System.EventHandler(this.chkCode_CheckedChanged);
            // 
            // chkAllTypes
            // 
            this.chkAllTypes.AutoSize = true;
            this.chkAllTypes.Location = new System.Drawing.Point(24, 23);
            this.chkAllTypes.Name = "chkAllTypes";
            this.chkAllTypes.Size = new System.Drawing.Size(73, 17);
            this.chkAllTypes.TabIndex = 10;
            this.chkAllTypes.Text = "Все типы";
            this.chkAllTypes.UseVisualStyleBackColor = true;
            this.chkAllTypes.CheckedChanged += new System.EventHandler(this.chkAllTypes_CheckedChanged);
            // 
            // chkComplement
            // 
            this.chkComplement.AutoSize = true;
            this.chkComplement.Location = new System.Drawing.Point(24, 207);
            this.chkComplement.Name = "chkComplement";
            this.chkComplement.Size = new System.Drawing.Size(84, 17);
            this.chkComplement.TabIndex = 7;
            this.chkComplement.Text = "Комплекты";
            this.chkComplement.UseVisualStyleBackColor = true;
            // 
            // chkMaterial
            // 
            this.chkMaterial.AutoSize = true;
            this.chkMaterial.Location = new System.Drawing.Point(24, 184);
            this.chkMaterial.Name = "chkMaterial";
            this.chkMaterial.Size = new System.Drawing.Size(84, 17);
            this.chkMaterial.TabIndex = 6;
            this.chkMaterial.Text = "Материалы";
            this.chkMaterial.UseVisualStyleBackColor = true;
            // 
            // chkOther
            // 
            this.chkOther.AutoSize = true;
            this.chkOther.Location = new System.Drawing.Point(24, 161);
            this.chkOther.Name = "chkOther";
            this.chkOther.Size = new System.Drawing.Size(108, 17);
            this.chkOther.TabIndex = 5;
            this.chkOther.Text = "Прочие изделия";
            this.chkOther.UseVisualStyleBackColor = true;
            // 
            // chkStandart
            // 
            this.chkStandart.AutoSize = true;
            this.chkStandart.Location = new System.Drawing.Point(24, 138);
            this.chkStandart.Name = "chkStandart";
            this.chkStandart.Size = new System.Drawing.Size(138, 17);
            this.chkStandart.TabIndex = 4;
            this.chkStandart.Text = "Стандартные изделия";
            this.chkStandart.UseVisualStyleBackColor = true;
            // 
            // chkAssembly
            // 
            this.chkAssembly.AutoSize = true;
            this.chkAssembly.Location = new System.Drawing.Point(24, 69);
            this.chkAssembly.Name = "chkAssembly";
            this.chkAssembly.Size = new System.Drawing.Size(129, 17);
            this.chkAssembly.TabIndex = 3;
            this.chkAssembly.Text = "Сборочные единицы";
            this.chkAssembly.ThreeState = true;
            this.chkAssembly.UseVisualStyleBackColor = true;
            this.chkAssembly.CheckedChanged += new System.EventHandler(this.chkAssembly_CheckedChanged);
            // 
            // chkDetail
            // 
            this.chkDetail.AutoSize = true;
            this.chkDetail.Location = new System.Drawing.Point(24, 46);
            this.chkDetail.Name = "chkDetail";
            this.chkDetail.Size = new System.Drawing.Size(64, 17);
            this.chkDetail.TabIndex = 2;
            this.chkDetail.Text = "Детали";
            this.chkDetail.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(770, 286);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 3;
            // 
            // btnMakeReport
            // 
            this.btnMakeReport.Enabled = false;
            this.btnMakeReport.Location = new System.Drawing.Point(564, 500);
            this.btnMakeReport.Name = "btnMakeReport";
            this.btnMakeReport.Size = new System.Drawing.Size(178, 23);
            this.btnMakeReport.TabIndex = 4;
            this.btnMakeReport.Text = "Сформировать отчет";
            this.btnMakeReport.UseVisualStyleBackColor = true;
            this.btnMakeReport.Click += new System.EventHandler(this.btnMakeReport_Click);
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(767, 185);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(182, 23);
            this.buttonClose.TabIndex = 5;
            this.buttonClose.Text = "Выход";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // dgvProducts
            // 
            this.dgvProducts.AllowDrop = true;
            this.dgvProducts.AllowUserToAddRows = false;
            this.dgvProducts.AllowUserToDeleteRows = false;
            this.dgvProducts.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader;
            this.dgvProducts.Location = new System.Drawing.Point(16, 41);
            this.dgvProducts.MultiSelect = false;
            this.dgvProducts.Name = "dgvProducts";
            this.dgvProducts.ReadOnly = true;
            this.dgvProducts.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvProducts.Size = new System.Drawing.Size(726, 243);
            this.dgvProducts.TabIndex = 6;
            // 
            // dgvAddedProducts
            // 
            this.dgvAddedProducts.AllowDrop = true;
            this.dgvAddedProducts.AllowUserToAddRows = false;
            this.dgvAddedProducts.AllowUserToDeleteRows = false;
            this.dgvAddedProducts.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvAddedProducts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAddedProducts.Location = new System.Drawing.Point(16, 332);
            this.dgvAddedProducts.MultiSelect = false;
            this.dgvAddedProducts.Name = "dgvAddedProducts";
            this.dgvAddedProducts.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvAddedProducts.ShowEditingIcon = false;
            this.dgvAddedProducts.Size = new System.Drawing.Size(726, 156);
            this.dgvAddedProducts.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(166, 13);
            this.label2.TabIndex = 14;
            this.label2.Text = "Список заказов для обработки";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 303);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(159, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "Список добавленных заказов";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(16, 555);
            this.progressBar1.Maximum = 100000;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(726, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 20;
            // 
            // progressLabel
            // 
            this.progressLabel.AllowDrop = true;
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(18, 539);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(109, 17);
            this.progressLabel.TabIndex = 19;
            this.progressLabel.Text = "Добавление заказа";
            this.progressLabel.UseCompatibleTextRendering = true;
            // 
            // deleteComplexButton
            // 
            this.deleteComplexButton.Enabled = false;
            this.deleteComplexButton.Location = new System.Drawing.Point(16, 500);
            this.deleteComplexButton.Name = "deleteComplexButton";
            this.deleteComplexButton.Size = new System.Drawing.Size(109, 23);
            this.deleteComplexButton.TabIndex = 18;
            this.deleteComplexButton.Text = "Удалить расчет";
            this.deleteComplexButton.UseVisualStyleBackColor = true;
            this.deleteComplexButton.Click += new System.EventHandler(this.deleteComplexButton_Click);
            // 
            // clearDBButton
            // 
            this.clearDBButton.Enabled = false;
            this.clearDBButton.Location = new System.Drawing.Point(141, 500);
            this.clearDBButton.Name = "clearDBButton";
            this.clearDBButton.Size = new System.Drawing.Size(118, 23);
            this.clearDBButton.TabIndex = 17;
            this.clearDBButton.Text = "Удалить все";
            this.clearDBButton.UseVisualStyleBackColor = true;
            this.clearDBButton.Click += new System.EventHandler(this.clearDBButton_Click);
            // 
            // addProductButton
            // 
            this.addProductButton.Location = new System.Drawing.Point(764, 156);
            this.addProductButton.Name = "addProductButton";
            this.addProductButton.Size = new System.Drawing.Size(182, 23);
            this.addProductButton.TabIndex = 21;
            this.addProductButton.Text = "Добавить заказ";
            this.addProductButton.UseVisualStyleBackColor = true;
            this.addProductButton.Click += new System.EventHandler(this.addProductButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBoxComplement);
            this.groupBox1.Controls.Add(this.checkBoxComplex);
            this.groupBox1.Location = new System.Drawing.Point(764, 60);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(182, 73);
            this.groupBox1.TabIndex = 23;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Тип заказов";
            // 
            // checkBoxComplement
            // 
            this.checkBoxComplement.AutoSize = true;
            this.checkBoxComplement.Location = new System.Drawing.Point(20, 42);
            this.checkBoxComplement.Name = "checkBoxComplement";
            this.checkBoxComplement.Size = new System.Drawing.Size(84, 17);
            this.checkBoxComplement.TabIndex = 1;
            this.checkBoxComplement.Text = "Комплекты";
            this.checkBoxComplement.UseVisualStyleBackColor = true;
            this.checkBoxComplement.CheckedChanged += new System.EventHandler(this.checkBoxComplement_CheckedChanged);
            // 
            // checkBoxComplex
            // 
            this.checkBoxComplex.AutoSize = true;
            this.checkBoxComplex.Checked = true;
            this.checkBoxComplex.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxComplex.Location = new System.Drawing.Point(20, 19);
            this.checkBoxComplex.Name = "checkBoxComplex";
            this.checkBoxComplex.Size = new System.Drawing.Size(85, 17);
            this.checkBoxComplex.TabIndex = 0;
            this.checkBoxComplex.Text = "Комплексы";
            this.checkBoxComplex.UseVisualStyleBackColor = true;
            this.checkBoxComplex.CheckedChanged += new System.EventHandler(this.checkBoxComplex_CheckedChanged);
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(961, 589);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.addProductButton);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.deleteComplexButton);
            this.Controls.Add(this.clearDBButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dgvAddedProducts);
            this.Controls.Add(this.dgvProducts);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.btnMakeReport);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет - расчет общего количества изделий в комплексе";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProducts)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAddedProducts)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkAllTypes;
        private System.Windows.Forms.CheckBox chkComplement;
        private System.Windows.Forms.CheckBox chkMaterial;
        private System.Windows.Forms.CheckBox chkOther;
        private System.Windows.Forms.CheckBox chkStandart;
        private System.Windows.Forms.CheckBox chkAssembly;
        private System.Windows.Forms.CheckBox chkDetail;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnMakeReport;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.DataGridView dgvProducts;
        private System.Windows.Forms.DataGridView dgvAddedProducts;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label progressLabel;
        private System.Windows.Forms.Button deleteComplexButton;
        private System.Windows.Forms.Button clearDBButton;
        private System.Windows.Forms.Button addProductButton;
        private System.Windows.Forms.CheckBox chkCodeless;
        private System.Windows.Forms.CheckBox chkCode;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBoxComplement;
        private System.Windows.Forms.CheckBox checkBoxComplex;
    }
}