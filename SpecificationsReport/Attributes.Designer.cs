using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SpecificationsReport.Vspecnew;


using TFlex.Reporting;
//using SpecificationESKD.Globus.TFlexDocs;


namespace SpecificationsReport
{
    partial class Attributes
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
            this.makeReportButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.nudReservePages = new System.Windows.Forms.NumericUpDown();
            this.grBoxModifySpec.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudReservePages)).BeginInit();
            this.SuspendLayout();
            // 
            // chBoxAddModify
            // 
            this.chBoxAddModify.AutoSize = true;
            this.chBoxAddModify.Location = new System.Drawing.Point(196, 17);
            this.chBoxAddModify.Name = "chBoxAddModify";
            this.chBoxAddModify.Size = new System.Drawing.Size(135, 17);
            this.chBoxAddModify.TabIndex = 9;
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
            this.grBoxModifySpec.Location = new System.Drawing.Point(20, 52);
            this.grBoxModifySpec.Name = "grBoxModifySpec";
            this.grBoxModifySpec.Size = new System.Drawing.Size(507, 225);
            this.grBoxModifySpec.TabIndex = 8;
            this.grBoxModifySpec.TabStop = false;
            this.grBoxModifySpec.Text = "Листы и информация об изменених";
            this.grBoxModifySpec.Enter += new System.EventHandler(this.grBoxModifySpec_Enter);
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
            // makeReportButton
            // 
            this.makeReportButton.Location = new System.Drawing.Point(420, 345);
            this.makeReportButton.Name = "makeReportButton";
            this.makeReportButton.Size = new System.Drawing.Size(107, 23);
            this.makeReportButton.TabIndex = 10;
            this.makeReportButton.Text = "Создать отчет";
            this.makeReportButton.UseVisualStyleBackColor = true;
            this.makeReportButton.Click += new System.EventHandler(this.makeReportButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.nudReservePages);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(20, 283);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(507, 52);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Резервные страницы";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(112, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Количество";
            // 
            // nudReservePages
            // 
            this.nudReservePages.Location = new System.Drawing.Point(194, 20);
            this.nudReservePages.Name = "nudReservePages";
            this.nudReservePages.Size = new System.Drawing.Size(68, 20);
            this.nudReservePages.TabIndex = 1;
            // 
            // Attributes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(549, 378);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.makeReportButton);
            this.Controls.Add(this.chBoxAddModify);
            this.Controls.Add(this.grBoxModifySpec);
            this.Name = "Attributes";
            this.Text = "Настройка отчета";
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Attributes_FormClosed);
            this.Load += new System.EventHandler(this.Attributes_Load);
            this.grBoxModifySpec.ResumeLayout(false);
            this.grBoxModifySpec.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudReservePages)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion


        private bool _attributesFormClose = false;
        private bool _closeGenerator = false;
        private bool closeForm = false;

        private void makeReportButton_Click(object sender, EventArgs e)
        {
            _attributesFormClose = true;
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
        public void DotsEntry(IReportGenerationContext context,   MainForm m_form, LogDelegate m_writeToLog)
        {
            this.Show();

            while (AttributesFormClose == false)
            {
                this.UpdateUI();
                if (closeForm)
                {
                    m_form.Close();
                    break;

                }
            }
            if (closeForm)
            {
                this.Close();
                return;
            }

            if (CloseGenerator)
            {
                _closeGenerator = false;
                this.Close();
                return;
            }

            DocumentAttributes documentAttr = new DocumentAttributes();

            if (chBoxAddModify.CheckState == CheckState.Checked)
                documentAttr.AddModify = true;
            else documentAttr.AddModify = false;

            documentAttr.countReservPages = Convert.ToInt32(nudReservePages.Value);

            if (chBoxAddModify.CheckState == CheckState.Checked)
            {
                
                documentAttr.NumberModify = textBoxNumberModify.Text;
                documentAttr.NumberDocument = textBoxNumberDoc.Text;
                documentAttr.DateModify = textBoxDate.Text;
                documentAttr.ListZam = textBoxList.Text;
                documentAttr.ListNov = textBoxList2.Text;
                

                int index = 0;
                int indexDash = 0;
                string str = textBoxListPage.Text.Trim();
                string strWithDash = string.Empty;
                int firstPage = 0;
                int secondPage = 0;

                // Преобразование строки с номерами страниц через запятую в список номеров страниц
                while (str != string.Empty)
                {
                    index = str.IndexOf(",");

                    if (index != -1)
                    {
                        strWithDash = str.Substring(0, index).Trim();
                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {

                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = str.Remove(0, str.Substring(0, index - indexDash).Length);

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage.Add(i);

                        }
                        else
                        {

                            documentAttr.ListPage.Add(Convert.ToInt32(str.Substring(0, index).Trim()));
                            str = str.Remove(0, str.Substring(0, index + 1).Length);
                        }
                    }
                    else
                    {

                        index = str.Length;
                        strWithDash = str.Trim();



                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {
                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = string.Empty;

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage.Add(i);

                        }
                        else
                        {
                            documentAttr.ListPage.Add(Convert.ToInt32(str.Trim()));
                            str = string.Empty;
                        }
                    }

                }

                index = 0;
                indexDash = 0;
                str = textBoxListPage2.Text.Trim();
                strWithDash = string.Empty;
                firstPage = 0;
                secondPage = 0;


                // Преобразование строки с номерами страниц через запятую в список номеров страниц listPage2
                while (str != string.Empty)
                {
                    index = str.IndexOf(",");

                    if (index != -1)
                    {
                        strWithDash = str.Substring(0, index).Trim();
                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {

                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = str.Remove(0, str.Substring(0, index - indexDash).Length);

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage2.Add(i);

                        }
                        else
                        {

                            documentAttr.ListPage2.Add(Convert.ToInt32(str.Substring(0, index).Trim()));
                            str = str.Remove(0, str.Substring(0, index + 1).Length);
                        }
                    }
                    else
                    {

                        index = str.Length;
                        strWithDash = str.Trim();



                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {
                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = string.Empty;

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage2.Add(i);

                        }
                        else
                        {
                            documentAttr.ListPage2.Add(Convert.ToInt32(str.Trim()));
                            str = string.Empty;
                        }
                    }

                }


            }

            this.Close();

            // ТОЧКА ВХОДА МАКРОСА
            m_form.Visible = true;
            Specification.MakeSpecListReport(context, documentAttr, m_form, m_writeToLog);
           
        }


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
        private System.Windows.Forms.Button makeReportButton;
        private GroupBox groupBox1;
        private NumericUpDown nudReservePages;
        private Label label1;
    }
}