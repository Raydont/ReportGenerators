namespace GroupingOIReport
{
    partial class NotFoundObjectForm
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
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonAddObjects = new System.Windows.Forms.Button();
            this.comboBoxSections = new System.Windows.Forms.ComboBox();
            this.buttonMoving = new System.Windows.Forms.Button();
            this.listViewNotFoundObjects = new System.Windows.Forms.ListView();
            this.Number = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.NameObject = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.DenotationObject = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Parent = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(811, 349);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(92, 23);
            this.buttonCancel.TabIndex = 22;
            this.buttonCancel.Text = "Отмена";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonAddObjects
            // 
            this.buttonAddObjects.Location = new System.Drawing.Point(24, 349);
            this.buttonAddObjects.Name = "buttonAddObjects";
            this.buttonAddObjects.Size = new System.Drawing.Size(781, 23);
            this.buttonAddObjects.TabIndex = 21;
            this.buttonAddObjects.Text = "Добавить объекты в отчет";
            this.buttonAddObjects.UseVisualStyleBackColor = true;
            this.buttonAddObjects.Click += new System.EventHandler(this.buttonAddObjects_Click);
            // 
            // comboBoxSections
            // 
            this.comboBoxSections.FormattingEnabled = true;
            this.comboBoxSections.Location = new System.Drawing.Point(749, 170);
            this.comboBoxSections.Name = "comboBoxSections";
            this.comboBoxSections.Size = new System.Drawing.Size(154, 21);
            this.comboBoxSections.TabIndex = 20;
            // 
            // buttonMoving
            // 
            this.buttonMoving.Enabled = false;
            this.buttonMoving.Location = new System.Drawing.Point(706, 170);
            this.buttonMoving.Name = "buttonMoving";
            this.buttonMoving.Size = new System.Drawing.Size(37, 23);
            this.buttonMoving.TabIndex = 19;
            this.buttonMoving.Text = ">";
            this.buttonMoving.UseVisualStyleBackColor = true;
            this.buttonMoving.Click += new System.EventHandler(this.buttonMoving_Click);
            // 
            // listViewNotFoundObjects
            // 
            this.listViewNotFoundObjects.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Number,
            this.NameObject,
            this.DenotationObject,
            this.Parent});
            this.listViewNotFoundObjects.FullRowSelect = true;
            this.listViewNotFoundObjects.GridLines = true;
            this.listViewNotFoundObjects.Location = new System.Drawing.Point(24, 26);
            this.listViewNotFoundObjects.Name = "listViewNotFoundObjects";
            this.listViewNotFoundObjects.Size = new System.Drawing.Size(667, 308);
            this.listViewNotFoundObjects.TabIndex = 18;
            this.listViewNotFoundObjects.UseCompatibleStateImageBehavior = false;
            this.listViewNotFoundObjects.View = System.Windows.Forms.View.Details;
            this.listViewNotFoundObjects.SelectedIndexChanged += new System.EventHandler(this.listViewNotFoundObjects_SelectedIndexChanged);
            // 
            // Number
            // 
            this.Number.Text = "№";
            this.Number.Width = 30;
            // 
            // NameObject
            // 
            this.NameObject.Text = "Наименование";
            this.NameObject.Width = 296;
            // 
            // DenotationObject
            // 
            this.DenotationObject.Text = "Обозначение";
            this.DenotationObject.Width = 109;
            // 
            // Parent
            // 
            this.Parent.Text = "Устройство";
            this.Parent.Width = 205;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(746, 149);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 23;
            this.label1.Text = "Выбор группы";
            // 
            // NotFoundObjectForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(915, 389);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonAddObjects);
            this.Controls.Add(this.comboBoxSections);
            this.Controls.Add(this.buttonMoving);
            this.Controls.Add(this.listViewNotFoundObjects);
            this.Name = "NotFoundObjectForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ненайденные объекты";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonAddObjects;
        public System.Windows.Forms.ComboBox comboBoxSections;
        private System.Windows.Forms.Button buttonMoving;
        public System.Windows.Forms.ListView listViewNotFoundObjects;
        private System.Windows.Forms.ColumnHeader Number;
        private System.Windows.Forms.ColumnHeader NameObject;
        private System.Windows.Forms.ColumnHeader DenotationObject;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ColumnHeader Parent;
    }
}