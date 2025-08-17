namespace PeriodicTestsChart
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
            this.buttonChangeFolder = new System.Windows.Forms.Button();
            this.textBoxPathFolder = new System.Windows.Forms.TextBox();
            this.radioButtonLoadedOrder = new System.Windows.Forms.RadioButton();
            this.radioButtonAllOrders = new System.Windows.Forms.RadioButton();
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBoxNomenclatureReference.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBoxNomenclatureReference
            // 
            this.groupBoxNomenclatureReference.Controls.Add(this.buttonChangeFolder);
            this.groupBoxNomenclatureReference.Controls.Add(this.textBoxPathFolder);
            this.groupBoxNomenclatureReference.Controls.Add(this.radioButtonLoadedOrder);
            this.groupBoxNomenclatureReference.Controls.Add(this.radioButtonAllOrders);
            this.groupBoxNomenclatureReference.Location = new System.Drawing.Point(12, 22);
            this.groupBoxNomenclatureReference.Name = "groupBoxNomenclatureReference";
            this.groupBoxNomenclatureReference.Size = new System.Drawing.Size(546, 100);
            this.groupBoxNomenclatureReference.TabIndex = 0;
            this.groupBoxNomenclatureReference.TabStop = false;
            this.groupBoxNomenclatureReference.Text = "Номенклатурные справочники";
            // 
            // buttonChangeFolder
            // 
            this.buttonChangeFolder.Enabled = false;
            this.buttonChangeFolder.Location = new System.Drawing.Point(482, 55);
            this.buttonChangeFolder.Name = "buttonChangeFolder";
            this.buttonChangeFolder.Size = new System.Drawing.Size(44, 23);
            this.buttonChangeFolder.TabIndex = 9;
            this.buttonChangeFolder.Text = "...";
            this.buttonChangeFolder.UseVisualStyleBackColor = true;
            this.buttonChangeFolder.Click += new System.EventHandler(this.buttonChangeFolder_Click_1);
            // 
            // textBoxPathFolder
            // 
            this.textBoxPathFolder.Enabled = false;
            this.textBoxPathFolder.Location = new System.Drawing.Point(194, 57);
            this.textBoxPathFolder.Name = "textBoxPathFolder";
            this.textBoxPathFolder.Size = new System.Drawing.Size(296, 20);
            this.textBoxPathFolder.TabIndex = 7;
            this.textBoxPathFolder.TextChanged += new System.EventHandler(this.textBoxPathFolder_TextChanged_1);
            // 
            // radioButtonLoadedOrder
            // 
            this.radioButtonLoadedOrder.AutoSize = true;
            this.radioButtonLoadedOrder.Location = new System.Drawing.Point(194, 27);
            this.radioButtonLoadedOrder.Name = "radioButtonLoadedOrder";
            this.radioButtonLoadedOrder.Size = new System.Drawing.Size(303, 17);
            this.radioButtonLoadedOrder.TabIndex = 2;
            this.radioButtonLoadedOrder.Text = "Отчет по загруженному списку устройств или заказов";
            this.radioButtonLoadedOrder.UseVisualStyleBackColor = true;
            this.radioButtonLoadedOrder.CheckedChanged += new System.EventHandler(this.radioButtonLoadedOrder_CheckedChanged);
            // 
            // radioButtonAllOrders
            // 
            this.radioButtonAllOrders.AutoSize = true;
            this.radioButtonAllOrders.Checked = true;
            this.radioButtonAllOrders.Location = new System.Drawing.Point(17, 27);
            this.radioButtonAllOrders.Name = "radioButtonAllOrders";
            this.radioButtonAllOrders.Size = new System.Drawing.Size(145, 17);
            this.radioButtonAllOrders.TabIndex = 0;
            this.radioButtonAllOrders.TabStop = true;
            this.radioButtonAllOrders.Text = "Отчет по всем заказам";
            this.radioButtonAllOrders.UseVisualStyleBackColor = true;
            this.radioButtonAllOrders.CheckedChanged += new System.EventHandler(this.radioButtonAllOrders_CheckedChanged);
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(12, 137);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(546, 23);
            this.buttonMakeReport.TabIndex = 1;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click_1);
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.HelpRequest += new System.EventHandler(this.folderBrowserDialog1_HelpRequest);
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(568, 179);
            this.Controls.Add(this.buttonMakeReport);
            this.Controls.Add(this.groupBoxNomenclatureReference);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчёт для составления графика ПИ";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBoxNomenclatureReference.ResumeLayout(false);
            this.groupBoxNomenclatureReference.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxNomenclatureReference;
        private System.Windows.Forms.RadioButton radioButtonLoadedOrder;
        private System.Windows.Forms.RadioButton radioButtonAllOrders;
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button buttonChangeFolder;
        private System.Windows.Forms.TextBox textBoxPathFolder;
    }
}