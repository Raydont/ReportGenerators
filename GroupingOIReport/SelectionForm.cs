using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;

namespace GroupingOIReport
{
    public partial class SelectionForm : Form
    {
        private IReportGenerationContext _context;
        public SelectionForm()
        {
            InitializeComponent();
        }
        public ReportParameters reportParameters = new ReportParameters();

        public void Init(IReportGenerationContext context)
        {
            _context = context;
            SetSettingsForm();
        }

        private void SetSettingsForm()
        {
            var referenceInfo =  ReferenceCatalog.FindReference(_context.ReferenceId);
            var reference = referenceInfo.CreateReference();

            if (reference.Name == "Регистрационно-контрольные карточки")
            {
                label2.Visible = false;
                textBoxFileName1.Visible = false;
                buttonLoadFile1.Visible = false;
                groupBoxOptions.Location = new System.Drawing.Point(21, 19);
                var order = GetSelectedObject();
                var lastFile = order.GetObjects(Guids.RegistartionCards.СвязьДокументы).OrderBy(t => t.SystemFields.EditDate).Last();
                if (lastFile == null)
                {
                    MessageBox.Show("Внимание!", $"У объекта {order.ToString()} отсутствуют прикрепленные файлы. Отчет не будет сформирован", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    return;
                }
                var file = (FileReferenceObject)lastFile;
                file.GetHeadRevision();
                loadExcelData(11, 3, 2, 5, file.LocalPath);
                textBoxOrderNumber.Text = order[Guids.RegistartionCards.ПараметрНомерЗаказа].ToString();
                textBoxSpecificationOrder.Text = order.ToString();
            }
            else 
            {
                label2.Visible = true;
                textBoxFileName1.Visible = true;
                buttonLoadFile1.Visible = true;
                groupBoxOptions.Location = new System.Drawing.Point(21, 51);
            }
        }

        private ReferenceObject GetSelectedObject()
        {
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(_context.ReferenceId);
            var reference = referenceInfo.CreateReference();

            reference.LoadSettings.AddAllParameters();
            return reference.Find(_context.ObjectsInfo.First().ObjectId);       
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            GroupsReferenceNames.Initializer();
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            if (MakeReportParameters())
                DialogResult = DialogResult.OK;
        }

        private bool MakeReportParameters()
        {
            reportParameters = new ReportParameters();

            reportParameters.StandartItemsType = checkBoxStandartItems.CheckState == CheckState.Checked ? true : false;
            reportParameters.OtherItems = checkBoxOtherItems.CheckState == CheckState.Checked ? true : false;
            reportParameters.DetailsItems = checkBoxDetails.CheckState == CheckState.Checked ? true : false;

            reportParameters.AddCodeDevices = checkBoxAddCodeDevices.CheckState == CheckState.Checked ? true : false;
            reportParameters.AddAllCodeDevice = checkBoxAddAllCodeObject.CheckState == CheckState.Checked ? true : false;
            reportParameters.NotCodeDevices = checkBoxNotCodeDevices.CheckState == CheckState.Checked ? true : false;
            reportParameters.IsResultManyFiles = checkBoxIsResultManyFiles.CheckState == CheckState.Checked ? true : false;
            reportParameters.IsCreateObjectCost = checkBoxIsCreateObjectCost.CheckState == CheckState.Checked ? true : false;
            reportParameters.OrderNumber = textBoxOrderNumber.Text;
            reportParameters.SpecificationOrder = textBoxSpecificationOrder.Text;

            if (nomObjects.Count == 0)
            {
                MessageBox.Show("Данные не удалось считать. Исходные данные должны начинаться со второй строки в формате: \r\n" +
                   "| 1колонка | 2колонка           | 3колонка           | 4колонка     | 5колонка     |\r\n"+
                   "| № п.п.      | Шифр        | Наименование | Обозначение | Количество |", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            foreach (var obj in nomObjects)
            {
                reportParameters.NomObjects.Add(obj);
            }

            return true;
        }

        public List<NomSpecificationItems> nomObjects = new List<NomSpecificationItems>();

        private void buttonLoadFile1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "(Excel-файлы)|*.xls;*.xlsx";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxFileName1.Text = openFileDialog1.FileName;
                if (!textBoxFileName1.Text.ToLower().Contains(".xls"))
                    return;

                loadExcelData(2, 4, 3, 5, openFileDialog1.FileName);
            }
            else
            {
                return;
            }
        }

        private void loadExcelData(int beginRow, int denotationColumn, int nameColumn, int counColumn, string filePath)
        {
            var errorXlsObjects = new List<string>();
            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            int errorID = 0;
            int numberStr = 0;
            var xls = new Xls();
            
            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(filePath, false);
                errorID = 1;
                var reference = nomenclatureReferenceInfo.CreateReference();
                errorID = 2;
                for (int i = beginRow; xls[denotationColumn, i].Text.Trim() != string.Empty; i++)
                {
                    numberStr = i;
                    errorID = 3;
                    var denotation = xls[denotationColumn, i].Text.Trim();
                    var name = xls[nameColumn, i].Text.Trim();
                   // var number = xls[1, i].Text.Trim();
                    errorID = 4;
                    int count = Convert.ToInt32(xls[counColumn, i].Text.Trim());
                    errorID = 5;

                    var filter = new Filter(nomenclatureReferenceInfo);
                    filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
                    errorID = 6;
                    var documentObject = reference.Find(filter).FirstOrDefault();
                    errorID = 7;
                    if (documentObject != null)
                    {
                        var newObj = new NomSpecificationItems((NomenclatureObject)documentObject, name, count);

                        var existObject = nomObjects.Where(t => t.NomObject == newObj.NomObject).ToList();

                        if (existObject.Count() == 0 || existObject.Count(t => t.NameLoadList != newObj.NameLoadList) > 0)
                            nomObjects.Add(newObj);
                        else
                            nomObjects.Where(t => t.NomObject == newObj.NomObject).FirstOrDefault().Count += newObj.Count;

                        errorID = 8;
                    }
                    else
                    {
                        errorID = 9;
                        errorXlsObjects.Add(denotation + " строка № " + i);
                    }
                }

                if (errorXlsObjects.Count > 0)
                {
                    var notFoundObjToStr = string.Join("\r\n", errorXlsObjects);
                    Clipboard.SetText(notFoundObjToStr);
                    MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("2. Внимание ошибка! Сформированный отчет будет не корректен. Обратитесь в отдел 911. Возникла ошибка при работе с файлом в строке " + numberStr + "\r\n" + filePath + "\r\nКод ошибки:" + errorID + "\r\n" + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                xls.Quit(true);
                xls.Close();
            }
            buttonMakeReport.Enabled = true;
        }

        private void checkBoxFirstLevel_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxNotCodeDevices.Checked)
            {
                checkBoxAddAllCodeObject.Checked = false;
                checkBoxAddCodeDevices.Checked = false;
            }   
        }

        private void checkBoxAddCodeDevices_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAddCodeDevices.Checked)
            {
                checkBoxNotCodeDevices.Checked = false;
            }

            if (!checkBoxAddCodeDevices.Checked && checkBoxAddAllCodeObject.Checked)
            {
                checkBoxAddAllCodeObject.Checked = false;
            }
        }

        private void checkBoxAddAllCodeObject_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAddAllCodeObject.Checked)
            {
                checkBoxNotCodeDevices.Checked = false;
                checkBoxAddCodeDevices.Checked = true;
            }
        }
    }
}
