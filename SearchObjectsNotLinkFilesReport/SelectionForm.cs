using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.UI.Common;
using TFlex.DOCs.UI.Common.Filters;
using TFlex.Reporting;

namespace SearchObjectsNotLinkFilesReport
{
    public partial class SelectionForm : Form
    {
        public static Guid CodeGroupGuid = new Guid("bef7a0c1-3d3a-48c9-a9ef-9e5c51cd3be4");
        public static Guid CodeGuid = new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6");
        public static Guid NomenclatureReferenceGuid = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        public static Guid TypeComplexGuid = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");



        private IReportGenerationContext _context;

        public SelectionForm()
        {
            InitializeComponent();
        }

        public void Init(IReportGenerationContext context)
        {

            _context = context;

        }

        private void buttonDeleteOrder_Click(object sender, EventArgs e)
        {
            listViewOrder.Items.Remove(listViewOrder.FocusedItem);
            buttonDeleteOrder.Enabled = false;
        }

        Filter filter;


        private void buttonAddOrder_Click(object sender, EventArgs e)
        {

 

            using (var filterDialog = new ReferenceFilterDialog())
            {
                var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(NomenclatureReferenceGuid);
                if (filter != null)
                {
                    filterDialog.Initialize(filter);
                }
                else
                {
                    filter = new Filter(nomenclatureReferenceInfo);
                    filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, "БФМИ.461271.");
                    filterDialog.Initialize(filter);
                    
                }
                try
                {

                if (filterDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    filter = filterDialog.Filter;

                    var nomenclatureReference = nomenclatureReferenceInfo.CreateReference();

                    var objects = nomenclatureReference.Find(filterDialog.Filter);


                    //Вывод наименований
                    foreach (NomenclatureObject nomObject in objects)
                    {
                        if (nomObject.Class.IsInherit(TypeComplexGuid))
                        {
                            var item = new ListViewItem(nomObject.SystemFields.Id.ToString());
                            item.SubItems.Add(nomObject.Denotation);
                            item.SubItems.Add(nomObject[CodeGuid].ToString());
                            item.Tag = nomObject;
                            listViewOrder.Items.Add(item);
                        }
                        else
                        {
                            MessageBox.Show("Объект " + nomObject.Denotation + " не является заказом!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        } 
                    }
                }

                }
                catch (Exception)
                {

                    MessageBox.Show("Ошибка! Обратитесь в бюро 911." + e.ToString(),"Внимание!",MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

       

        private void listViewOrder_SelectedIndexChanged(object sender, EventArgs e)
        {
            var lv = sender as ListView;
            if ((listViewOrder.Items.Count != 0) && (listViewOrder.SelectedItems.Count != 0) && (listViewOrder.FocusedItem.Selected))
            {
                buttonDeleteOrder.Enabled = true;
            }
            else
            {
                buttonDeleteOrder.Enabled = false;
            }
        }

        public List<int> listOrder = new List<int>()
                                         {120133,262125,302338,326934,245584,245581};

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(NomenclatureReferenceGuid);
            var nomenclatureReference = nomenclatureReferenceInfo.CreateReference();
            foreach (var orderId in listOrder)
            {
                var nomObject = (NomenclatureObject)nomenclatureReference.Find(orderId);
                var item = new ListViewItem(nomObject.SystemFields.Id.ToString());
                item.SubItems.Add(nomObject.Denotation);
                item.SubItems.Add(nomObject[CodeGuid].ToString());
                item.Tag = nomObject;
                listViewOrder.Items.Add(item);
            }          
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            MakeReportParameters();
            DialogResult = DialogResult.OK;
        }

        public AttributeParametrsClass attributeParameters;

        private void MakeReportParameters()
        {
            attributeParameters = new AttributeParametrsClass();
            var makeReport = false;

            while (!makeReport)
            {
                makeReport = true;

                if (checkBoxDocument.CheckState == CheckState.Checked)
                    attributeParameters.DocumentObject = true;
                else
                    attributeParameters.DocumentObject = false;

                if (checkBoxDetail.CheckState == CheckState.Checked)
                    attributeParameters.DetailObject = true;
                else
                    attributeParameters.DetailObject = false;

                if (checkBoxComplement.CheckState == CheckState.Checked)
                    attributeParameters.ComplementObject = true;
                else
                    attributeParameters.ComplementObject = false;

                if (comboBoxTypeReport.Text == "Объекты с неприсоединенными файлами")
                    attributeParameters.typeReport = true;
                else
                    attributeParameters.typeReport = false;



                foreach (ListViewItem itmx in listViewOrder.Items)
                {
                    attributeParameters.ListOrderId.Add(Convert.ToInt32(itmx.SubItems[0].Text.Trim()));
                }

                if (listViewOrder.Items.Count == 0)
                {
                    MessageBox.Show("Не выбранно ни одного заказа! Выберите заказ для отчета!", "Внимание!",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    makeReport = false;
                }
               

                if (attributeParameters.DocumentObject != true && attributeParameters.DetailObject != true && attributeParameters.ComplementObject != true)
                {
                    MessageBox.Show("Не выбранно ни одного типа объектов для отчета!", "Внимание!",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    makeReport = false;
                }
            }

        }
    }
}
