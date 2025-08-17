using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;


namespace PeriodicTestsChart
{
    public partial class SelectionForm : Form
    {
        private IReportGenerationContext _context;
        public ReportParameters reportParameters = new ReportParameters();
        public SelectionForm()
        {
            InitializeComponent();
        }
        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        private void buttonMakeReport_Click_1(object sender, EventArgs e)
        {
            reportParameters.AllOrder = radioButtonAllOrders.Checked;
            reportParameters.LoadedListOrder = radioButtonLoadedOrder.Checked;
            if (radioButtonLoadedOrder.Checked && reportParameters.OrdersList.Count == 0)
            {
                MessageBox.Show("В списке загруженных заказов отсутствуют данные. Выберите файл со списком заказов", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult = DialogResult.OK;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
       
            textBoxPathFolder.Text = "";
        }

        private void textBoxPathFolder_TextChanged(object sender, EventArgs e)
        {


        }


        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void buttonChangeFolder_Click(object sender, EventArgs e)
        {
          
        }
        private static Filter CreateFilter(ReferenceInfo documentReferenceInfo, string denotation)
        {

            var filter = new Filter(documentReferenceInfo);
            filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
            return filter;
        }

        private void radioButtonAllOrders_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonAllOrders.Checked)
            {
                buttonChangeFolder.Enabled = false;
                textBoxPathFolder.Enabled = false;
            }
        }

        private void radioButtonLoadedOrder_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonLoadedOrder.Checked)
            {
                buttonChangeFolder.Enabled = true;
                textBoxPathFolder.Enabled = true;
            }
        }

        private void buttonChangeFolder_Click_1(object sender, EventArgs e)
        {
            var xls = new Xls();
            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    xls.Application.DisplayAlerts = false;
               
                    xls.Open(fileDialog.FileName, false);

                    var reference = nomenclatureReferenceInfo.CreateReference();

                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty; i++)
                    {
                     //   MessageBox.Show(fileDialog.FileName);
                        var denotation = xls[1, i].Text.Trim();

                        var filter = CreateFilter(nomenclatureReferenceInfo, denotation);


                       List<ReferenceObject> documentObjects = reference.Find(filter);
                      //  MessageBox.Show(string.Join("\r\n", documentObjects.Select(t=>t.ToString())));

                        if (documentObjects == null)
                            MessageBox.Show("Объект " + xls[1, i].Value + " не найден в базе");
                        else
                        {
                            reportParameters.OrdersList.Add((NomenclatureObject)documentObjects.First());
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Возникла ошибка при работе с файлом " + fileDialog.FileName + "\n" + ex);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
            }

        }

        private void textBoxPathFolder_TextChanged_1(object sender, EventArgs e)
        {

        }
    }



    public class ReportParameters
    {
        public bool AllOrder = false;
        public bool LoadedListOrder = false;
        public List<NomenclatureObject> OrdersList = new List<NomenclatureObject>();

    }
}
