using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using TFlex.Reporting;

namespace BuyProductsReport
{
    public partial class SelectionParamsForm : Form
    {
        public SelectionParamsForm()
        {
            InitializeComponent();
        }

        private bool _selectionFormClose = false;
        private bool _closeGenerator = false;


        void UpdateUI()
        {
            Update();
            Application.DoEvents();
        }
        
        public bool SelectionFormClose
        {
            get { return _selectionFormClose; }
        }

        public bool CloseGenerator
        {
            get { return _closeGenerator; }
        }

        public void DotsEntry(IReportGenerationContext context)
        {
            this.Show();

            DocumentAttributes documentAttr = new DocumentAttributes();

            while (SelectionFormClose == false)
            {
                this.UpdateUI();
            }

            if (CloseGenerator)
            {
                _closeGenerator = false;
                this.Close();
                return;
            }

            if (chBoxAddList.CheckState == CheckState.Checked) documentAttr.AddList = true;
            else documentAttr.AddList = false;

            if (chBoxSectionNewList.CheckState == CheckState.Checked) documentAttr.NewListSection = true;
            else documentAttr.NewListSection = false;


            if (chBoxAddModify.CheckState == CheckState.Checked)
            {
                documentAttr.AddModify = true;
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
            BuyListReport.MakeBuyListReport(context, documentAttr);
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            _selectionFormClose = true;
        }

        private void SelectionParamsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _selectionFormClose = true;
            _closeGenerator = true;
        }

        private void chBoxSectionNewList_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chBoxAddModify_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAddModify.Checked)
                grBoxModifySpec.Enabled = true;
            else grBoxModifySpec.Enabled = false;

        }
    }
}
