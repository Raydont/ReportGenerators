using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using System.Text.RegularExpressions;

namespace PreciosMetalsReport
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
            if (listNames.Count > 0)
                reportParameters.PKIData = reportParameters.PKIData.Where(t => listNames.Count(n=>t.Name.Contains(n)) > 0).ToList();
            if (reportParameters.PKIData.Count == 0)
            {
                MessageBox.Show("В списке загруженных объектов ПКИ отсутствуют данные.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

         //   Clipboard.SetText(string.Join("\r\n", reportParameters.PKIData.Select(t=>t.StoreObject.ToString() + " " + t.RemainsEndCount)));
           // MessageBox.Show("Считанные данные");

            DialogResult = DialogResult.OK;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {

            textBoxLoadFile.Text = "";
        }


        public string AddStoreObject(Reference reference, ReferenceInfo documentReferenceInfo, int row, object[,] range)
        {
            var rowError = string.Empty;
            if (range[row, 1].ToString().Trim() == string.Empty)
                return string.Empty;
            var filter = new Filter(documentReferenceInfo);

            var namePKI = range[row, 2].ToString().Trim();

          //  if (listNames.Count == 0 || listNames.Count(t => namePKI.Trim().Contains(t)) > 0)
           // {
                var storeObject = new StoreData(namePKI,
                       range[row, 3].ToString().Trim(),
                       range[row, 4].ToString().Trim(),
                       range[row, 5].ToString().Trim(),
                       range[row, 6].ToString().Trim());
            if (!string.IsNullOrEmpty(storeObject.Error))
            {
                return storeObject.Error;
            }

                filter.Terms.AddTerm("Артикул", ComparisonOperator.Equal, storeObject.Article);

                List<ReferenceObject> nomStoreObjects = reference.Find(filter);

            if (nomStoreObjects == null || nomStoreObjects.Count == 0 || storeObject.Article == string.Empty)
                rowError = namePKI + " Строка № " + row;
            else
            {
                storeObject.StoreObject = nomStoreObjects.First();

                if (Convert.ToDouble(storeObject.StoreObject[Guids.NomStore.GoldContains].Value) > 0 ||
            Convert.ToDouble(storeObject.StoreObject[Guids.NomStore.SilverContains].Value) > 0 ||
            Convert.ToDouble(storeObject.StoreObject[Guids.NomStore.PlatinaContains].Value) > 0)
                    reportParameters.PKIData.Add(storeObject);
                // }
            }
            return rowError;
        }


        private void buttonChangeFolder_Click_1(object sender, EventArgs e)
        {
            var xls = new Xls();
            var nomStoreReferenceInfo = ReferenceCatalog.FindReference(Guids.NomenclatureStoreGuid);
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxLoadFile.Text = fileDialog.FileName;
                var errorID = 0;
                var row = 0;
                xls.Application.DisplayAlerts = false;
                errorID = 2;
                xls.Open(fileDialog.FileName, false);
                var range = (object[,])xls[1, 1, 6, 16000].Value2;

                try
                {
                    var reference = nomStoreReferenceInfo.CreateReference();
                   
                    var strBError = new List <string>();
                    
                    for (int i = 2; range[i, 1] != null; i++)
                    {
                        row = i;
                        errorID = 2551;
                        strBError.Add(AddStoreObject(reference, nomStoreReferenceInfo, i, range));
                        errorID = 255;
                    }
                    if (strBError.Count > 0)
                    {
                        var strBErrorWithOutEmpty = strBError.Where(t => t != string.Empty);

                        if (strBErrorWithOutEmpty.Count() > 0)
                        {
                            Clipboard.SetText(string.Join("\r\n", strBErrorWithOutEmpty));
                            MessageBox.Show("Не найдены следующие объекты:\r\n" + string.Join("\r\n", strBErrorWithOutEmpty));
                        }
                    }
                    errorID = 5;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Возникла ошибка при работе с файлом. Код ошибки " + errorID + " |" + range[row,1] + " |" + row + "| " + fileDialog.FileName + "\n" + ex);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
            }

            buttonMakeReport.Enabled = true;
       

        }

        public List<string> listNames = new List<string>();

        private void addFilterButton_Click(object sender, EventArgs e)
        {
            var filterText = textBoxAddLikeName.Text.Trim();

            if (filterText != string.Empty)
            {
                var itmx = new ListViewItem[1];
                itmx[0] = new ListViewItem((filterListView.Items.Count + 1).ToString());
                itmx[0].SubItems.Add(filterText);
                itmx[0].Tag = filterText;
                listNames.Add(filterText);

                filterListView.Items.AddRange(itmx);
            }
        }

        private void buttonClearList_Click(object sender, EventArgs e)
        {
            filterListView.Items.Clear();
            listNames = new List<string>();
        }
    }

}
