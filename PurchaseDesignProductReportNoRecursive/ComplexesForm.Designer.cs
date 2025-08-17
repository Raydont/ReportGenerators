using System.Windows.Forms;
using TFlex.Reporting;

namespace PurchaseDesignProductReportForm
{
    partial class ComplexesForm
    {
        private bool _attributesFormClose = false;
        private bool _closeGenerator = false;

        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        void UpdateUI()
        {
            Update();
            Application.DoEvents();
        }

        public bool AttributesFormClose
        {
            get { return _attributesFormClose; }
        }

        public bool CloseGenerator
        {
            get { return _closeGenerator; }
        }

        public void DotsEntry(string reportFilePath)
        {
            reportFilePathInForm = reportFilePath;
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgvProducts = new System.Windows.Forms.DataGridView();
            this.dgvAddedProducts = new System.Windows.Forms.DataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.addProductButton = new System.Windows.Forms.Button();
            this.clearDBButton = new System.Windows.Forms.Button();
            this.deleteComplexButton = new System.Windows.Forms.Button();
            this.reportButton = new System.Windows.Forms.Button();
            this.reportPanel = new System.Windows.Forms.Panel();
            this.rbGroupStart = new System.Windows.Forms.RadioButton();
            this.rbPurchaseProducts = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.progressLabel = new System.Windows.Forms.Label();
            this.cbFirstUnitCodeDevices = new System.Windows.Forms.CheckBox();
            this.cbOccurence = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvProducts)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAddedProducts)).BeginInit();
            this.reportPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvProducts
            // 
            this.dgvProducts.AllowUserToAddRows = false;
            this.dgvProducts.AllowUserToDeleteRows = false;
            this.dgvProducts.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvProducts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvProducts.Location = new System.Drawing.Point(21, 75);
            this.dgvProducts.MultiSelect = false;
            this.dgvProducts.Name = "dgvProducts";
            this.dgvProducts.ReadOnly = true;
            this.dgvProducts.Size = new System.Drawing.Size(726, 306);
            this.dgvProducts.TabIndex = 1;
            this.dgvProducts.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvProducts_CellClick);
            // 
            // dgvAddedProducts
            // 
            this.dgvAddedProducts.AllowUserToAddRows = false;
            this.dgvAddedProducts.AllowUserToDeleteRows = false;
            this.dgvAddedProducts.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvAddedProducts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvAddedProducts.Location = new System.Drawing.Point(21, 405);
            this.dgvAddedProducts.MultiSelect = false;
            this.dgvAddedProducts.Name = "dgvAddedProducts";
            this.dgvAddedProducts.ShowEditingIcon = false;
            this.dgvAddedProducts.Size = new System.Drawing.Size(726, 150);
            this.dgvAddedProducts.TabIndex = 2;
            this.dgvAddedProducts.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvAddedProducts_CellClick);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 389);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(159, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Список добавленных заказов";
            // 
            // addProductButton
            // 
            this.addProductButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.addProductButton.Location = new System.Drawing.Point(638, 12);
            this.addProductButton.Name = "addProductButton";
            this.addProductButton.Size = new System.Drawing.Size(109, 23);
            this.addProductButton.TabIndex = 4;
            this.addProductButton.Text = "Добавить заказ";
            this.addProductButton.UseVisualStyleBackColor = true;
            this.addProductButton.Click += new System.EventHandler(this.addProductButton_Click);
            // 
            // clearDBButton
            // 
            this.clearDBButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.clearDBButton.Location = new System.Drawing.Point(146, 570);
            this.clearDBButton.Name = "clearDBButton";
            this.clearDBButton.Size = new System.Drawing.Size(118, 23);
            this.clearDBButton.TabIndex = 5;
            this.clearDBButton.Text = "Удалить все";
            this.clearDBButton.UseVisualStyleBackColor = true;
            this.clearDBButton.Click += new System.EventHandler(this.clearDBButton_Click);
            // 
            // deleteComplexButton
            // 
            this.deleteComplexButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.deleteComplexButton.Location = new System.Drawing.Point(21, 570);
            this.deleteComplexButton.Name = "deleteComplexButton";
            this.deleteComplexButton.Size = new System.Drawing.Size(109, 23);
            this.deleteComplexButton.TabIndex = 7;
            this.deleteComplexButton.Text = "Удалить расчет";
            this.deleteComplexButton.UseVisualStyleBackColor = true;
            this.deleteComplexButton.Click += new System.EventHandler(this.deleteComplexButton_Click);
            // 
            // reportButton
            // 
            this.reportButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.reportButton.Location = new System.Drawing.Point(657, 570);
            this.reportButton.Name = "reportButton";
            this.reportButton.Size = new System.Drawing.Size(90, 23);
            this.reportButton.TabIndex = 8;
            this.reportButton.Text = "Отчет";
            this.reportButton.UseVisualStyleBackColor = true;
            this.reportButton.Click += new System.EventHandler(this.reportButton_Click);
            // 
            // reportPanel
            // 
            this.reportPanel.Controls.Add(this.rbGroupStart);
            this.reportPanel.Controls.Add(this.rbPurchaseProducts);
            this.reportPanel.Location = new System.Drawing.Point(21, 12);
            this.reportPanel.Name = "reportPanel";
            this.reportPanel.Size = new System.Drawing.Size(598, 23);
            this.reportPanel.TabIndex = 12;
            // 
            // rbGroupStart
            // 
            this.rbGroupStart.AutoSize = true;
            this.rbGroupStart.Location = new System.Drawing.Point(309, 3);
            this.rbGroupStart.Name = "rbGroupStart";
            this.rbGroupStart.Size = new System.Drawing.Size(185, 17);
            this.rbGroupStart.TabIndex = 10;
            this.rbGroupStart.Text = "Ведомость группового запуска";
            this.rbGroupStart.UseVisualStyleBackColor = true;
            this.rbGroupStart.CheckedChanged += new System.EventHandler(this.rbGroupStart_CheckedChanged);
            // 
            // rbPurchaseProducts
            // 
            this.rbPurchaseProducts.AutoSize = true;
            this.rbPurchaseProducts.Checked = true;
            this.rbPurchaseProducts.Location = new System.Drawing.Point(28, 3);
            this.rbPurchaseProducts.Name = "rbPurchaseProducts";
            this.rbPurchaseProducts.Size = new System.Drawing.Size(266, 17);
            this.rbPurchaseProducts.TabIndex = 0;
            this.rbPurchaseProducts.TabStop = true;
            this.rbPurchaseProducts.Text = "Ведомость покупных конструкторских изделий";
            this.rbPurchaseProducts.UseVisualStyleBackColor = true;
            this.rbPurchaseProducts.CheckedChanged += new System.EventHandler(this.rbPurchaseProducts_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(166, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Список заказов для обработки";
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(21, 625);
            this.progressBar1.Maximum = 4;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(726, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 15;
            // 
            // progressLabel
            // 
            this.progressLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(18, 607);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(109, 13);
            this.progressLabel.TabIndex = 14;
            this.progressLabel.Text = "Добавление заказа";
            // 
            // cbFirstUnitCodeDevices
            // 
            this.cbFirstUnitCodeDevices.AutoSize = true;
            this.cbFirstUnitCodeDevices.Location = new System.Drawing.Point(330, 38);
            this.cbFirstUnitCodeDevices.Name = "cbFirstUnitCodeDevices";
            this.cbFirstUnitCodeDevices.Size = new System.Drawing.Size(333, 17);
            this.cbFirstUnitCodeDevices.TabIndex = 16;
            this.cbFirstUnitCodeDevices.Text = "с первыми шифрованными (только для группового запуска)";
            this.cbFirstUnitCodeDevices.UseVisualStyleBackColor = true;
            // 
            // cbOccurence
            // 
            this.cbOccurence.AutoSize = true;
            this.cbOccurence.Checked = true;
            this.cbOccurence.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbOccurence.Enabled = false;
            this.cbOccurence.Location = new System.Drawing.Point(49, 38);
            this.cbOccurence.Name = "cbOccurence";
            this.cbOccurence.Size = new System.Drawing.Size(192, 17);
            this.cbOccurence.TabIndex = 19;
            this.cbOccurence.Text = "с указанием первой входимости";
            this.cbOccurence.UseVisualStyleBackColor = true;
            // 
            // ComplexesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 660);
            this.Controls.Add(this.cbOccurence);
            this.Controls.Add(this.cbFirstUnitCodeDevices);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.reportPanel);
            this.Controls.Add(this.reportButton);
            this.Controls.Add(this.deleteComplexButton);
            this.Controls.Add(this.clearDBButton);
            this.Controls.Add(this.addProductButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dgvAddedProducts);
            this.Controls.Add(this.dgvProducts);
            this.MinimumSize = new System.Drawing.Size(783, 699);
            this.Name = "ComplexesForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Список заказов для обработки";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ComplexesForm_FormClosed);
            this.Load += new System.EventHandler(this.ComplexesForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvProducts)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvAddedProducts)).EndInit();
            this.reportPanel.ResumeLayout(false);
            this.reportPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvProducts;
        private System.Windows.Forms.DataGridView dgvAddedProducts;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button addProductButton;
        private System.Windows.Forms.Button clearDBButton;
        private System.Windows.Forms.Button deleteComplexButton;
        private System.Windows.Forms.Button reportButton;
        private System.Windows.Forms.Panel reportPanel;
        private System.Windows.Forms.RadioButton rbGroupStart;
        private System.Windows.Forms.RadioButton rbPurchaseProducts;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label progressLabel;
        private CheckBox cbFirstUnitCodeDevices;
        private CheckBox cbOccurence;
    }
}

