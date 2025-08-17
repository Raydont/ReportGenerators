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
using TFlex.DOCs.UI.Common;
using TFlex.DOCs.UI.Common.Filters;
using TFlex.Reporting;

namespace TermsPeriodicTestsReport
{
    public partial class SelectionForm : Form
    {
        private IReportGenerationContext _context;
        public ReportParameters reportParameters;

        public void Init(IReportGenerationContext context)
        {
            _context = context;
            

            ReferenceInfo referenceInfo = ReferenceCatalog.FindReference(context.ReferenceId);
            Reference reference = referenceInfo.CreateReference();

            if (reference.ToString() == "Номенклатура и изделия")
            {
                var refObject = reference.Find(_context.ObjectsInfo[0].ObjectId);

                listViewFindObjects.Items.Clear();

                var item = new ListViewItem(refObject.SystemFields.Id.ToString());
                item.SubItems.Add(refObject[Guids.Nomenclature.Name].Value.ToString());
                item.SubItems.Add(refObject[Guids.Nomenclature.Denotation].Value.ToString());
                item.Tag = refObject;
                listViewFindObjects.Items.Add(item);
            }
        }

        public SelectionForm()
        {
            InitializeComponent();
        }

        Filter filter;

        //Вызов системного поиска объектов Номенклатуры
        private void buttonFindObject_Click(object sender, EventArgs e)
        {
            using (var filterDialog = new ReferenceFilterDialog())
            {
                var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(Guids.References.Nomenclature);

                if (filter != null)
                {
                    filterDialog.Initialize(filter);
                }
                else
                {
                    filterDialog.Initialize(nomenclatureReferenceInfo);
                }

                if (filterDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    filter = filterDialog.Filter;
                    var nomenclatureReference = nomenclatureReferenceInfo.CreateReference();

                    var objects = nomenclatureReference.Find(filterDialog.Filter);
                    listViewFindObjects.Items.Clear();
                    foreach (NomenclatureObject nomObject in objects)
                    {
                        var item = new ListViewItem(nomObject.SystemFields.Id.ToString());
                        item.SubItems.Add(nomObject.Name);
                        item.SubItems.Add(nomObject.Denotation);
                        item.Tag = nomObject;
                        listViewFindObjects.Items.Add(item);
                    }
                }
            }
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            if (listViewFindObjects.FocusedItem == null)
            {
                MessageBox.Show("Вы не выбрали ни одного объекта для отчета", "Внимание!!!", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }

            reportParameters = new ReportParameters();
            reportParameters.OrderName = (NomenclatureObject)listViewFindObjects.FocusedItem.Tag;
            reportParameters.TermDeliveryOrder = dateTimePickerOrder.Value.Date;

            DialogResult = DialogResult.OK;
        }

        private void listViewFindObjects_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewFindObjects.SelectedItems.Count > 0)
                buttonMakeReport.Enabled = true;
            else
                buttonMakeReport.Enabled = false;
        }

        private void SearchNameButton_Click(object sender, EventArgs e)
        {
            ReferenceInfo referenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            Reference reference = referenceInfo.CreateReference();

            //Создаем фильтр
            Filter filter = new Filter(referenceInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.IsOneOf, new[] { Guids.Nomenclature.ComplexType, Guids.Nomenclature.AssemblyType });    

            var termName  = filter.Terms.AddTerm("Наименование", ComparisonOperator.ContainsSubstring, textBoxNameOrDenotation.Text.Trim());
            var termDenotation = filter.Terms.AddTerm("Обозначение", ComparisonOperator.ContainsSubstring, textBoxNameOrDenotation.Text.Trim());
            termDenotation.LogicalOperator = LogicalOperator.Or;
            filter.Terms.GroupTerms(new[] { termName, termDenotation });
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр            
            List<ReferenceObject> listObj = reference.Find(filter);

            listViewFindObjects.Items.Clear();
            foreach (NomenclatureObject nomObject in listObj)
            {
                var item = new ListViewItem(nomObject.SystemFields.Id.ToString());
                item.SubItems.Add(nomObject.Name);
                item.SubItems.Add(nomObject.Denotation);
                item.Tag = nomObject;
                listViewFindObjects.Items.Add(item);
            }
        }
    }
}
