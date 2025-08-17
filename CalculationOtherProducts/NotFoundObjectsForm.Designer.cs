namespace CalculationOtherProducts
{
    partial class NotFoundObjectsForm
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
            this.listViewNotFoundObjects = new System.Windows.Forms.ListView();
            this.Number = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.NameObject = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.ShortName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.buttonMoving = new System.Windows.Forms.Button();
            this.comboBoxSections = new System.Windows.Forms.ComboBox();
            this.buttonAddObjects = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listViewNotFoundObjects
            // 
            this.listViewNotFoundObjects.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Number,
            this.NameObject,
            this.ShortName});
            this.listViewNotFoundObjects.FullRowSelect = true;
            this.listViewNotFoundObjects.GridLines = true;
            this.listViewNotFoundObjects.Location = new System.Drawing.Point(30, 26);
            this.listViewNotFoundObjects.Name = "listViewNotFoundObjects";
            this.listViewNotFoundObjects.Size = new System.Drawing.Size(671, 239);
            this.listViewNotFoundObjects.TabIndex = 13;
            this.listViewNotFoundObjects.UseCompatibleStateImageBehavior = false;
            this.listViewNotFoundObjects.View = System.Windows.Forms.View.Details;
            this.listViewNotFoundObjects.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.listViewNotFoundObjects_ItemCheck);
            this.listViewNotFoundObjects.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.listViewNotFoundObjects_ItemChecked);
            this.listViewNotFoundObjects.SelectedIndexChanged += new System.EventHandler(this.listViewNotFoundObjects_SelectedIndexChanged);
            this.listViewNotFoundObjects.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listViewNotFoundObjects_MouseDoubleClick);
            // 
            // Number
            // 
            this.Number.Text = "№";
            this.Number.Width = 30;
            // 
            // NameObject
            // 
            this.NameObject.Text = "Наименование";
            this.NameObject.Width = 380;
            // 
            // ShortName
            // 
            this.ShortName.Text = "Сокр. наименование";
            this.ShortName.Width = 230;
            // 
            // buttonMoving
            // 
            this.buttonMoving.Enabled = false;
            this.buttonMoving.Location = new System.Drawing.Point(727, 133);
            this.buttonMoving.Name = "buttonMoving";
            this.buttonMoving.Size = new System.Drawing.Size(37, 23);
            this.buttonMoving.TabIndex = 14;
            this.buttonMoving.Text = ">";
            this.buttonMoving.UseVisualStyleBackColor = true;
            this.buttonMoving.Click += new System.EventHandler(this.buttonMoving_Click);
            // 
            // comboBoxSections
            // 
            this.comboBoxSections.FormattingEnabled = true;
            this.comboBoxSections.Location = new System.Drawing.Point(794, 133);
            this.comboBoxSections.Name = "comboBoxSections";
            this.comboBoxSections.Size = new System.Drawing.Size(236, 21);
            this.comboBoxSections.TabIndex = 15;
            // 
            // buttonAddObjects
            // 
            this.buttonAddObjects.Location = new System.Drawing.Point(30, 284);
            this.buttonAddObjects.Name = "buttonAddObjects";
            this.buttonAddObjects.Size = new System.Drawing.Size(889, 23);
            this.buttonAddObjects.TabIndex = 16;
            this.buttonAddObjects.Text = "Добавить объекты в отчет";
            this.buttonAddObjects.UseVisualStyleBackColor = true;
            this.buttonAddObjects.Click += new System.EventHandler(this.buttonAddObjects_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(925, 284);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(105, 23);
            this.buttonCancel.TabIndex = 17;
            this.buttonCancel.Text = "Отмена";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // NotFoundObjectsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1042, 327);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonAddObjects);
            this.Controls.Add(this.comboBoxSections);
            this.Controls.Add(this.buttonMoving);
            this.Controls.Add(this.listViewNotFoundObjects);
            this.Name = "NotFoundObjectsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Объекты без категории";
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ListView listViewNotFoundObjects;
        private System.Windows.Forms.ColumnHeader Number;
        private System.Windows.Forms.ColumnHeader NameObject;
        private System.Windows.Forms.ColumnHeader ShortName;
        private System.Windows.Forms.Button buttonMoving;
        public System.Windows.Forms.ComboBox comboBoxSections;
        private System.Windows.Forms.Button buttonAddObjects;
        private System.Windows.Forms.Button buttonCancel;
    }
}