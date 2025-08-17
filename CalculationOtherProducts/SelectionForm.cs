using CalcualtionOtherProducts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace CalculationOtherProducts
{
    public partial class SelectionForm : Form
    {


        private IReportGenerationContext _context;
        public ReportParameters reportParameters;

        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        public SelectionForm()
        {
            InitializeComponent();
        }
        public ReportParameters MakeReport()
        {
            reportParameters = new ReportParameters();
            reportParameters.Assembly = (chkAssembly.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Details = (chkDetail.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Materials = (chkMaterial.CheckState == CheckState.Checked) ? true : false;
            reportParameters.OtherItems = (chkOther.CheckState == CheckState.Checked) ? true : false;
            reportParameters.StandartItems = (chkStandart.CheckState == CheckState.Checked) ? true : false;
            reportParameters.UseCurrentObject = (checkBoxUseCurrentObject.CheckState == CheckState.Checked) ? true : false;

            return reportParameters;
        }

        private void btnMakeReport_Click(object sender, EventArgs e)
        {
            if ((chkOther.CheckState == CheckState.Unchecked) && (chkAssembly.CheckState == CheckState.Unchecked) && (chkDetail.CheckState == CheckState.Unchecked) &&
              (chkStandart.CheckState == CheckState.Unchecked)&& (chkMaterial.CheckState == CheckState.Unchecked))
            {
                MessageBox.Show("Вы не выбрали ни одного типа объекта для отчета", "Внимание!!!", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }

            DialogResult = DialogResult.OK;
        }

        public List<NomSpecificationItems> devicesObjects = new List<NomSpecificationItems>();
        public List<NomSpecificationItems> pkiObjects = new List<NomSpecificationItems>();

        private void buttonLoadListDevice_Click(object sender, EventArgs e)
        {
            openFileDialogListDevice.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            List<string> errorXlsObjects = new List<string>();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            if (openFileDialogListDevice.ShowDialog() == DialogResult.OK)
            {
                var xls = new Xls();
                textBoxListDevice.Text = openFileDialogListDevice.FileName;
                if (!textBoxListDevice.Text.ToLower().Contains(".xls"))
                    return;

                int errorID = 0;

                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(openFileDialogListDevice.FileName, false);
                    errorID = 1;

                    var reference = nomenclatureReferenceInfo.CreateReference();
                    errorID = 2;

                    for (int i = 2; xls[4, i].Text.Trim() != string.Empty; i++)
                    {
                        errorID = 3;
                        var denotation = xls[4, i].Text.Trim();
                        var number = xls[1, i].Text.Trim();
                        errorID = 4;
                        int count = Convert.ToInt32(xls[5, i].Text.Trim());
                        string name = xls[3, i].Text.Trim();
                        errorID = 5;
                        var filter = new Filter(nomenclatureReferenceInfo);
                        filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
                        errorID = 6;
                        var documentObject = reference.Find(filter).FirstOrDefault();
                        errorID = 7;
                        if (documentObject != null)
                        {
                            errorID = 8;
                            var newObj = new NomSpecificationItems((NomenclatureObject)documentObject, name, count);

                            var existObject = devicesObjects.Where(t => t.NomObject == newObj.NomObject).ToList();

                            if (existObject.Count() == 0 || existObject.Count(t=>t.NameLoadList != newObj.NameLoadList) > 0)
                                devicesObjects.Add(newObj);
                            else
                                devicesObjects.Where(t => t.NomObject == newObj.NomObject).FirstOrDefault().Count += newObj.Count;
                        }
                        else
                        {
                            errorID = 9;
                            errorXlsObjects.Add(denotation + " строка № " + i);
                        }
                    }

                    if (errorXlsObjects.Count > 0)
                    {
                        MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("2. Возникла ошибка при работе с файлом " + openFileDialogListPKI.FileName + "\r\nКод ошибки:" + errorID + "\r\n" + ex);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
                btnMakeReport.Enabled = true;
            }
            else
                return;
        }

        private void buttonLoadListPKI_Click(object sender, EventArgs e)
        {
            openFileDialogListPKI.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            List<string> errorXlsObjects = new List<string>();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            if (openFileDialogListDevice.ShowDialog() == DialogResult.OK)
            {
                var xls = new Xls();
                textBoxListDevice.Text = openFileDialogListPKI.FileName;
                if (!textBoxListDevice.Text.ToLower().Contains(".xls"))
                    return;

                int errorID = 0;

                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(openFileDialogListPKI.FileName, false);
                    errorID = 1;

                    var reference = nomenclatureReferenceInfo.CreateReference();
                    errorID = 2;

                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty; i++)
                    {
                        errorID = 3;
                        var name = xls[1, i].Text.Trim();

                        errorID = 4;
                        int count = Convert.ToInt32(xls[2, i].Text.Trim());
                        errorID = 5;
                        var filter = new Filter(nomenclatureReferenceInfo);
                        filter.Terms.AddTerm("Наименование", ComparisonOperator.Equal, name);
                        errorID = 6;
                        var documentObject = reference.Find(filter).FirstOrDefault();
                        errorID = 7;
                        if (documentObject != null)
                        {
                            errorID = 8;
                            var newObj = new NomSpecificationItems((NomenclatureObject)documentObject, count);

                            if (pkiObjects.Count(t => t.NomObject == newObj.NomObject) == 0)
                                pkiObjects.Add(newObj);
                            else
                                pkiObjects.Where(t => t.NomObject == newObj.NomObject).FirstOrDefault().Count += newObj.Count;
                        }
                        else
                        {
                            errorID = 9;
                            errorXlsObjects.Add(name + " строка № " + i);
                        }
                    }

                    if (errorXlsObjects.Count > 0)
                    {
                        MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("2. Возникла ошибка при работе с файлом " + openFileDialogListDevice.FileName + "\r\nКод ошибки:" + errorID + "\r\n" + ex);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
                btnMakeReport.Enabled = true;
            }
            else
                return;
        }
    }
}
