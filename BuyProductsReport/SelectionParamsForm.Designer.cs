namespace BuyProductsReport
{
    partial class SelectionParamsForm
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
            this.chBoxAddList = new System.Windows.Forms.CheckBox();
            this.buttonMakeReport = new System.Windows.Forms.Button();
            this.chBoxSectionNewList = new System.Windows.Forms.CheckBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.chBoxAddModify = new System.Windows.Forms.CheckBox();
            this.grBoxModifySpec = new System.Windows.Forms.GroupBox();
            this.label25 = new System.Windows.Forms.Label();
            this.textBoxList2 = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.textBoxListPage2 = new System.Windows.Forms.TextBox();
            this.label24 = new System.Windows.Forms.Label();
            this.textBoxList = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.textBoxNumberModify = new System.Windows.Forms.TextBox();
            this.textBoxListPage = new System.Windows.Forms.TextBox();
            this.textBoxDate = new System.Windows.Forms.TextBox();
            this.textBoxNumberDoc = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.grBoxModifySpec.SuspendLayout();
            this.SuspendLayout();
            // 
            // chBoxAddList
            // 
            this.chBoxAddList.AutoSize = true;
            this.chBoxAddList.Location = new System.Drawing.Point(151, 44);
            this.chBoxAddList.Name = "chBoxAddList";
            this.chBoxAddList.Size = new System.Drawing.Size(200, 17);
            this.chBoxAddList.TabIndex = 0;
            this.chBoxAddList.Text = "Добавить лист после содержания";
            this.chBoxAddList.UseVisualStyleBackColor = true;
            // 
            // buttonMakeReport
            // 
            this.buttonMakeReport.Location = new System.Drawing.Point(336, 317);
            this.buttonMakeReport.Name = "buttonMakeReport";
            this.buttonMakeReport.Size = new System.Drawing.Size(207, 23);
            this.buttonMakeReport.TabIndex = 1;
            this.buttonMakeReport.Text = "Сформировать отчет";
            this.buttonMakeReport.UseVisualStyleBackColor = true;
            this.buttonMakeReport.Click += new System.EventHandler(this.buttonMakeReport_Click);
            // 
            // chBoxSectionNewList
            // 
            this.chBoxSectionNewList.AutoSize = true;
            this.chBoxSectionNewList.Location = new System.Drawing.Point(151, 73);
            this.chBoxSectionNewList.Name = "chBoxSectionNewList";
            this.chBoxSectionNewList.Size = new System.Drawing.Size(207, 17);
            this.chBoxSectionNewList.TabIndex = 2;
            this.chBoxSectionNewList.Text = "Начинать раздел с новой страницы";
            this.chBoxSectionNewList.UseVisualStyleBackColor = true;
            this.chBoxSectionNewList.CheckedChanged += new System.EventHandler(this.chBoxSectionNewList_CheckedChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(531, 299);
            this.tabControl1.TabIndex = 3;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.chBoxSectionNewList);
            this.tabPage1.Controls.Add(this.chBoxAddList);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(523, 273);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Настройка страниц";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.chBoxAddModify);
            this.tabPage2.Controls.Add(this.grBoxModifySpec);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(523, 273);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Добавить изменения";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // chBoxAddModify
            // 
            this.chBoxAddModify.AutoSize = true;
            this.chBoxAddModify.Location = new System.Drawing.Point(184, 15);
            this.chBoxAddModify.Name = "chBoxAddModify";
            this.chBoxAddModify.Size = new System.Drawing.Size(135, 17);
            this.chBoxAddModify.TabIndex = 12;
            this.chBoxAddModify.Text = "Добавить изменения";
            this.chBoxAddModify.UseVisualStyleBackColor = true;
            this.chBoxAddModify.CheckedChanged += new System.EventHandler(this.chBoxAddModify_CheckedChanged);
            // 
            // grBoxModifySpec
            // 
            this.grBoxModifySpec.Controls.Add(this.label25);
            this.grBoxModifySpec.Controls.Add(this.textBoxList2);
            this.grBoxModifySpec.Controls.Add(this.label26);
            this.grBoxModifySpec.Controls.Add(this.textBoxListPage2);
            this.grBoxModifySpec.Controls.Add(this.label24);
            this.grBoxModifySpec.Controls.Add(this.textBoxList);
            this.grBoxModifySpec.Controls.Add(this.label23);
            this.grBoxModifySpec.Controls.Add(this.label22);
            this.grBoxModifySpec.Controls.Add(this.label21);
            this.grBoxModifySpec.Controls.Add(this.label20);
            this.grBoxModifySpec.Controls.Add(this.textBoxNumberModify);
            this.grBoxModifySpec.Controls.Add(this.textBoxListPage);
            this.grBoxModifySpec.Controls.Add(this.textBoxDate);
            this.grBoxModifySpec.Controls.Add(this.textBoxNumberDoc);
            this.grBoxModifySpec.Enabled = false;
            this.grBoxModifySpec.Location = new System.Drawing.Point(8, 43);
            this.grBoxModifySpec.Name = "grBoxModifySpec";
            this.grBoxModifySpec.Size = new System.Drawing.Size(507, 225);
            this.grBoxModifySpec.TabIndex = 11;
            this.grBoxModifySpec.TabStop = false;
            this.grBoxModifySpec.Text = "Листы и информация об изменених";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(23, 164);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(32, 13);
            this.label25.TabIndex = 17;
            this.label25.Text = "Лист";
            // 
            // textBoxList2
            // 
            this.textBoxList2.Location = new System.Drawing.Point(194, 161);
            this.textBoxList2.Name = "textBoxList2";
            this.textBoxList2.Size = new System.Drawing.Size(152, 20);
            this.textBoxList2.TabIndex = 16;
            this.textBoxList2.Text = "Нов.";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(23, 191);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(165, 13);
            this.label26.TabIndex = 15;
            this.label26.Text = "Список листов (через запятую)";
            // 
            // textBoxListPage2
            // 
            this.textBoxListPage2.Location = new System.Drawing.Point(194, 188);
            this.textBoxListPage2.Name = "textBoxListPage2";
            this.textBoxListPage2.Size = new System.Drawing.Size(284, 20);
            this.textBoxListPage2.TabIndex = 14;
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(23, 109);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(32, 13);
            this.label24.TabIndex = 9;
            this.label24.Text = "Лист";
            // 
            // textBoxList
            // 
            this.textBoxList.Location = new System.Drawing.Point(194, 106);
            this.textBoxList.Name = "textBoxList";
            this.textBoxList.Size = new System.Drawing.Size(152, 20);
            this.textBoxList.TabIndex = 8;
            this.textBoxList.Text = "Зам.";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(23, 136);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(165, 13);
            this.label23.TabIndex = 7;
            this.label23.Text = "Список листов (через запятую)";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(23, 84);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(92, 13);
            this.label22.TabIndex = 6;
            this.label22.Text = "Дата изменения";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(23, 58);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(75, 13);
            this.label21.TabIndex = 5;
            this.label21.Text = "№ документа";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(23, 32);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(100, 13);
            this.label20.TabIndex = 4;
            this.label20.Text = "Номер изменения";
            // 
            // textBoxNumberModify
            // 
            this.textBoxNumberModify.Location = new System.Drawing.Point(194, 30);
            this.textBoxNumberModify.Name = "textBoxNumberModify";
            this.textBoxNumberModify.Size = new System.Drawing.Size(152, 20);
            this.textBoxNumberModify.TabIndex = 3;
            // 
            // textBoxListPage
            // 
            this.textBoxListPage.Location = new System.Drawing.Point(194, 133);
            this.textBoxListPage.Name = "textBoxListPage";
            this.textBoxListPage.Size = new System.Drawing.Size(284, 20);
            this.textBoxListPage.TabIndex = 0;
            // 
            // textBoxDate
            // 
            this.textBoxDate.Location = new System.Drawing.Point(194, 81);
            this.textBoxDate.Name = "textBoxDate";
            this.textBoxDate.Size = new System.Drawing.Size(152, 20);
            this.textBoxDate.TabIndex = 1;
            // 
            // textBoxNumberDoc
            // 
            this.textBoxNumberDoc.Location = new System.Drawing.Point(194, 55);
            this.textBoxNumberDoc.Name = "textBoxNumberDoc";
            this.textBoxNumberDoc.Size = new System.Drawing.Size(152, 20);
            this.textBoxNumberDoc.TabIndex = 2;
            // 
            // SelectionParamsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(550, 346);
            this.Controls.Add(this.buttonMakeReport);
            this.Controls.Add(this.tabControl1);
            this.Name = "SelectionParamsForm";
            this.Text = "Настройка отчёта ВПИ";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SelectionParamsForm_FormClosed);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.grBoxModifySpec.ResumeLayout(false);
            this.grBoxModifySpec.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckBox chBoxAddList;
        private System.Windows.Forms.Button buttonMakeReport;
        private System.Windows.Forms.CheckBox chBoxSectionNewList;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.CheckBox chBoxAddModify;
        private System.Windows.Forms.GroupBox grBoxModifySpec;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.TextBox textBoxList2;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.TextBox textBoxListPage2;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.TextBox textBoxList;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox textBoxNumberModify;
        private System.Windows.Forms.TextBox textBoxListPage;
        private System.Windows.Forms.TextBox textBoxDate;
        private System.Windows.Forms.TextBox textBoxNumberDoc;
    }
}