using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace ObjectUsageCodeDevicesReport
{
    public partial class SelectionForm : Form
    {
        public SelectionForm()
        {
            InitializeComponent();
        }

        public void Init(IReportGenerationContext context)
        {
            this.context = context;
            var nomenclatureReference = new NomenclatureReference();
            textBoxListNames.Text = string.Join("\r\n",context.ObjectsInfo.Select(t => (NomenclatureObject)nomenclatureReference.Find(t.ObjectId)).ToList());
           
        }

        private IReportGenerationContext context;
        List<ReferenceObject> pkiObjects = new List<ReferenceObject>();

        private void SearchDevicesButton_Click(object sender, EventArgs e)
        {
            //Список наименований ПКИ для поиска
            var names = textBoxListNames.Text.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Replace("\"", "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim())
                    .Distinct()
                    .ToList();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            var reference = nomenclatureReferenceInfo.CreateReference();

            var ненайденныеОбъекты = new List<string>();

            foreach (var name in names)
            {
                var filter = new TFlex.DOCs.Model.Search.Filter(nomenclatureReferenceInfo);
                filter.Terms.AddTerm("Наименование", ComparisonOperator.Equal, name);
                var findedObject = reference.Find(filter).FirstOrDefault();
                if (findedObject != null)
                {
                    pkiObjects.Add(findedObject);
                }
                else
                {
                    ненайденныеОбъекты.Add(name);
                }           
            }
             
            if (ненайденныеОбъекты.Count > 0)
            {
                MessageBox.Show("Не найдены следующие объкты:\r\n" + string.Join("\r\n", ненайденныеОбъекты) + "\r\nВозможно он некорректно заведен в Номенлатуре (пробелы в начале или конце Наименования)", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
           
            DialogResult = DialogResult.OK;    
        }

        public void MakeReport()
        {
            ReportParameters reportParameters = new ReportParameters();
            reportParameters.OnlyCodeDevicesNames = rBOnlyCodeDevicesName.Checked ? true : false;
            reportParameters.FullInformation = rBFullInformation.Checked ? true : false;

            if (pkiObjects.Count > 0)
            {
                reportParameters.PKIObjects.AddRange(pkiObjects);
                ObjectUsageCodeDevicesReport.MakeReport(context, reportParameters);
            }
            else
            {
                MessageBox.Show("Отсутствуют объекты для отчета! Отчет сформировать нельзя!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
