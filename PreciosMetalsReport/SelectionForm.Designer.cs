namespace PreciosMetalsReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SelectionForm));
            this.buttonLoadFile = new System.Windows.Forms.Button();
            this.textBoxLoadFile = new System.Windows.Forms.TextBox();
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.textBoxInfo = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.buttonClearList = new System.Windows.Forms.Button();
            this.addFilterButton = new System.Windows.Forms.Button();
            this.nameGroupBox = new System.Windows.Forms.GroupBox();
            this.textBoxAddLikeName = new System.Windows.Forms.TextBox();
            this.filterListView = new System.Windows.Forms.ListView();
            this.columnHeaderNumber = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox1.SuspendLayout();
            this.nameGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonLoadFile
            // 
            this.buttonLoadFile.Location = new System.Drawing.Point(516, 90);
            this.buttonLoadFile.Name = "buttonLoadFile";
            this.buttonLoadFile.Size = new System.Drawing.Size(42, 23);
            this.buttonLoadFile.TabIndex = 9;
            this.buttonLoadFile.Text = "...";
            this.buttonLoadFile.UseVisualStyleBackColor = true;
            this.buttonLoadFile.Click += new System.EventHandler(this.buttonChangeFolder_Click_1);
            // 
            // textBoxLoadFile
            // 
            this.textBoxLoadFile.Location = new System.Drawing.Point(14, 92);
            this.textBoxLoadFile.Name = "textBoxLoadFile";
            this.textBoxLoadFile.Size = new System.Drawing.Size(510, 20);
            this.textBoxLoadFile.TabIndex = 7;
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(14, 407);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(546, 23);
            this.buttonMakeReport.TabIndex = 1;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click_1);
            // 
            // textBoxInfo
            // 
            this.textBoxInfo.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.textBoxInfo.Location = new System.Drawing.Point(12, 20);
            this.textBoxInfo.Multiline = true;
            this.textBoxInfo.Name = "textBoxInfo";
            this.textBoxInfo.Size = new System.Drawing.Size(544, 62);
            this.textBoxInfo.TabIndex = 11;
            this.textBoxInfo.Text = resources.GetString("textBoxInfo.Text");
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.buttonClearList);
            this.groupBox1.Controls.Add(this.addFilterButton);
            this.groupBox1.Controls.Add(this.nameGroupBox);
            this.groupBox1.Controls.Add(this.filterListView);
            this.groupBox1.Location = new System.Drawing.Point(14, 130);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(544, 263);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Дополнительные условия";
            // 
            // buttonClearList
            // 
            this.buttonClearList.Location = new System.Drawing.Point(285, 221);
            this.buttonClearList.Name = "buttonClearList";
            this.buttonClearList.Size = new System.Drawing.Size(105, 23);
            this.buttonClearList.TabIndex = 26;
            this.buttonClearList.Text = "Очистить список";
            this.buttonClearList.UseVisualStyleBackColor = true;
            this.buttonClearList.Click += new System.EventHandler(this.buttonClearList_Click);
            // 
            // addFilterButton
            // 
            this.addFilterButton.Location = new System.Drawing.Point(396, 221);
            this.addFilterButton.Name = "addFilterButton";
            this.addFilterButton.Size = new System.Drawing.Size(114, 23);
            this.addFilterButton.TabIndex = 25;
            this.addFilterButton.Text = "Добавить";
            this.addFilterButton.UseVisualStyleBackColor = true;
            this.addFilterButton.Click += new System.EventHandler(this.addFilterButton_Click);
            // 
            // nameGroupBox
            // 
            this.nameGroupBox.Controls.Add(this.textBoxAddLikeName);
            this.nameGroupBox.Location = new System.Drawing.Point(16, 151);
            this.nameGroupBox.Name = "nameGroupBox";
            this.nameGroupBox.Size = new System.Drawing.Size(494, 55);
            this.nameGroupBox.TabIndex = 24;
            this.nameGroupBox.TabStop = false;
            this.nameGroupBox.Text = "Добавление условия содержит Наименование";
            // 
            // textBoxAddLikeName
            // 
            this.textBoxAddLikeName.Location = new System.Drawing.Point(40, 24);
            this.textBoxAddLikeName.Name = "textBoxAddLikeName";
            this.textBoxAddLikeName.Size = new System.Drawing.Size(448, 20);
            this.textBoxAddLikeName.TabIndex = 0;
            // 
            // filterListView
            // 
            this.filterListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderNumber,
            this.columnHeaderName});
            this.filterListView.FullRowSelect = true;
            this.filterListView.GridLines = true;
            this.filterListView.HideSelection = false;
            this.filterListView.Location = new System.Drawing.Point(16, 23);
            this.filterListView.Name = "filterListView";
            this.filterListView.Size = new System.Drawing.Size(494, 113);
            this.filterListView.TabIndex = 23;
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
            this.columnHeaderName.Width = 394;
            // 
            // SelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 437);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.textBoxInfo);
            this.Controls.Add(this.buttonLoadFile);
            this.Controls.Add(this.buttonMakeReport);
            this.Controls.Add(this.textBoxLoadFile);
            this.Name = "SelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчёт по драгметаллам";
            this.Load += new System.EventHandler(this.SelectionForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.nameGroupBox.ResumeLayout(false);
            this.nameGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button buttonLoadFile;
        private System.Windows.Forms.TextBox textBoxLoadFile;
        private System.Windows.Forms.TextBox textBoxInfo;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListView filterListView;
        private System.Windows.Forms.ColumnHeader columnHeaderNumber;
        private System.Windows.Forms.ColumnHeader columnHeaderName;
        private System.Windows.Forms.GroupBox nameGroupBox;
        private System.Windows.Forms.TextBox textBoxAddLikeName;
        private System.Windows.Forms.Button addFilterButton;
        private System.Windows.Forms.Button buttonClearList;
    }
}
