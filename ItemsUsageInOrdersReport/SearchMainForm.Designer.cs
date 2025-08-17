namespace ItemsUsageInOrdersReport
{
    public partial class SearchMainForm
    {
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

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.reportButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.excludeButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.itemsCount = new System.Windows.Forms.Label();
            this.searchedItemsListView = new System.Windows.Forms.ListView();
            this.label3 = new System.Windows.Forms.Label();
            this.itemsInList = new System.Windows.Forms.Label();
            this.selectAllButton = new System.Windows.Forms.Button();
            this.unselectButton = new System.Windows.Forms.Button();
            this.checkBoxAddCount = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // reportButton
            // 
            this.reportButton.Location = new System.Drawing.Point(494, 431);
            this.reportButton.Name = "reportButton";
            this.reportButton.Size = new System.Drawing.Size(146, 23);
            this.reportButton.TabIndex = 0;
            this.reportButton.Text = "Сформировать отчет";
            this.reportButton.UseVisualStyleBackColor = true;
            this.reportButton.Click += new System.EventHandler(this.reportButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(183, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Наименование изделий содержит:";
            // 
            // searchTextBox
            // 
            this.searchTextBox.Location = new System.Drawing.Point(214, 15);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(345, 20);
            this.searchTextBox.TabIndex = 2;
            this.searchTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.searchTextBox_KeyDown);
            // 
            // searchButton
            // 
            this.searchButton.Location = new System.Drawing.Point(565, 13);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(75, 23);
            this.searchButton.TabIndex = 3;
            this.searchButton.Text = "Найти";
            this.searchButton.UseVisualStyleBackColor = true;
            this.searchButton.Click += new System.EventHandler(this.searchButton_Click);
            // 
            // excludeButton
            // 
            this.excludeButton.Location = new System.Drawing.Point(221, 431);
            this.excludeButton.Name = "excludeButton";
            this.excludeButton.Size = new System.Drawing.Size(133, 23);
            this.excludeButton.TabIndex = 5;
            this.excludeButton.Text = "Удалить выделенное";
            this.excludeButton.UseVisualStyleBackColor = true;
            this.excludeButton.Click += new System.EventHandler(this.excludeButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Найдено изделий:";
            // 
            // itemsCount
            // 
            this.itemsCount.AutoSize = true;
            this.itemsCount.Location = new System.Drawing.Point(117, 40);
            this.itemsCount.Name = "itemsCount";
            this.itemsCount.Size = new System.Drawing.Size(10, 13);
            this.itemsCount.TabIndex = 8;
            this.itemsCount.Text = "-";
            // 
            // searchedItemsListView
            // 
            this.searchedItemsListView.CheckBoxes = true;
            this.searchedItemsListView.Location = new System.Drawing.Point(15, 61);
            this.searchedItemsListView.Name = "searchedItemsListView";
            this.searchedItemsListView.Size = new System.Drawing.Size(625, 335);
            this.searchedItemsListView.TabIndex = 9;
            this.searchedItemsListView.UseCompatibleStateImageBehavior = false;
            this.searchedItemsListView.View = System.Windows.Forms.View.List;
            this.searchedItemsListView.KeyDown += new System.Windows.Forms.KeyEventHandler(this.searchedItemsListView_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(422, 40);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "В списке:";
            // 
            // itemsInList
            // 
            this.itemsInList.AutoSize = true;
            this.itemsInList.Location = new System.Drawing.Point(484, 40);
            this.itemsInList.Name = "itemsInList";
            this.itemsInList.Size = new System.Drawing.Size(10, 13);
            this.itemsInList.TabIndex = 11;
            this.itemsInList.Text = "-";
            // 
            // selectAllButton
            // 
            this.selectAllButton.Location = new System.Drawing.Point(15, 431);
            this.selectAllButton.Name = "selectAllButton";
            this.selectAllButton.Size = new System.Drawing.Size(96, 23);
            this.selectAllButton.TabIndex = 12;
            this.selectAllButton.Text = "Выделить все";
            this.selectAllButton.UseVisualStyleBackColor = true;
            this.selectAllButton.Click += new System.EventHandler(this.selectAllButton_Click);
            // 
            // unselectButton
            // 
            this.unselectButton.Location = new System.Drawing.Point(120, 431);
            this.unselectButton.Name = "unselectButton";
            this.unselectButton.Size = new System.Drawing.Size(93, 23);
            this.unselectButton.TabIndex = 13;
            this.unselectButton.Text = "Сбросить все";
            this.unselectButton.UseVisualStyleBackColor = true;
            this.unselectButton.Click += new System.EventHandler(this.unselectButton_Click);
            // 
            // checkBoxAddCount
            // 
            this.checkBoxAddCount.AutoSize = true;
            this.checkBoxAddCount.Location = new System.Drawing.Point(15, 406);
            this.checkBoxAddCount.Name = "checkBoxAddCount";
            this.checkBoxAddCount.Size = new System.Drawing.Size(163, 17);
            this.checkBoxAddCount.TabIndex = 14;
            this.checkBoxAddCount.Text = "Выводить количество ПКИ";
            this.checkBoxAddCount.UseVisualStyleBackColor = true;
            // 
            // SearchMainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(653, 466);
            this.Controls.Add(this.checkBoxAddCount);
            this.Controls.Add(this.unselectButton);
            this.Controls.Add(this.selectAllButton);
            this.Controls.Add(this.itemsInList);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.searchedItemsListView);
            this.Controls.Add(this.itemsCount);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.excludeButton);
            this.Controls.Add(this.searchButton);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.reportButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "SearchMainForm";
            this.Text = "Поиск элементов и формирование отчета";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button reportButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Button searchButton;
        private System.Windows.Forms.Button excludeButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label itemsCount;
        private System.Windows.Forms.ListView searchedItemsListView;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label itemsInList;
        private System.Windows.Forms.Button selectAllButton;
        private System.Windows.Forms.Button unselectButton;
        private System.Windows.Forms.CheckBox checkBoxAddCount;
    }
}

