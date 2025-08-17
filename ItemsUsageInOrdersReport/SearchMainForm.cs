using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReportHelpers;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

namespace ItemsUsageInOrdersReport
{
    public partial class SearchMainForm : Form
    {
        SqlConnection conn;
        string fileName;

        // справочник "Номенклатура и изделия"
        private readonly Guid NomenclatureReferenceGuid = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        // тип "Прочее изделие" справочника "Номенклатура и изделия"
        private readonly Guid OtherItemClassGuid = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
        // параметр "Наименование" справочника "Номенклатура и изделия"
        private readonly Guid NameParameterGuid = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");

        // соединение с DOCs
        public static SqlConnection GetConnection()
        {
            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            var info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            var reference = info.CreateReference();
            var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
            sqlConStringBuilder.DataSource = nameSqlServer;
            sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
            sqlConStringBuilder.Password = "reportUser";
            sqlConStringBuilder.UserID = "reportUser";
            return new SqlConnection(sqlConStringBuilder.ToString());
        }

        public SearchMainForm(string filePath)
        {
            fileName = filePath;
            InitializeComponent();
            conn = GetConnection();
            conn.Open();
        }

        private void reportButton_Click(object sender, EventArgs e)
        {
            var text = Text;
            Enabled = false;

            DateTime startTime = DateTime.Now;
            ReferenceInfo referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.Nomenclature);
            // создание объекта для работы с данными
            Reference reference = referenceInfo.CreateReference();

  

            var itemsUsages = new Dictionary<int, List<List<TFDDocument>>>();
            var items = searchedItemsListView.Items
                .Cast<ListViewItem>()
                .Where(t=>t.Checked)
                .Select(t => (NomenclatureObject)t.Tag)
                .OrderBy(t => t.Name.GetString())
                .ToList();

            if (items.Count == 0)
            {
                MessageBox.Show("Для отчета ничего не выбрано");
                return;
            }


            for (int i = 0; i < items.Count; i++)
            {
                NomenclatureObject item = items[i];
                Text = "Получение информации " + (i + 1) + " из " + items.Count + " " + (DateTime.Now - startTime);
                Application.DoEvents();
                var itemId = item.SystemFields.Id;
                var itemOrderBranches = new ItemOrderBranches(itemId);
                var branches = itemOrderBranches.SelectionOrderBranch(conn, true);
                itemsUsages[itemId] = branches;

            }

            var total = itemsUsages.Select(t => t.Value.Count).Sum();


            using (var xls = new Xls())
            {
                int additionalCol = 0;
                int row = 1;
                xls[1, row].SetValue("Ведомость применяемости ПКИ");
                xls[1, row].CenterText();
                xls[1, row, 3, 1].Merge();
                row++;
                row++;
                var tableStart = row;
                xls[1, row].SetValue("Наименвоание и обозначение");

                if (checkBoxAddCount.Checked)
                {
                    xls[2, row].SetValue("Количество");
                    additionalCol = 1;
                }

                xls[2 + additionalCol, row].SetValue("Входимость");
                xls[3 + additionalCol, row].SetValue("Заказ");
                xls[1, row, 3 + additionalCol, 1].CenterText();
                row++;

                var start = row;
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    var itemUsages = itemsUsages[item.SystemFields.Id];

                    Text = "Вывод информации " + (i + 1) + " из " + items.Count + " [строка " + (row - start + 1) + " из " + total + "] " + (DateTime.Now - startTime);
                    Application.DoEvents();

                    var first = true;
                    foreach (var itemUsage in itemUsages)
                    {
                        if (first)
                        {
                            xls[1, row].SetValue((item.Name + " " + item.Denotation).Trim());
                            first = false;
                        }

                        if (checkBoxAddCount.Checked)
                        {
                            var parent = (NomenclatureObject)reference.Find(itemUsage[1].Id);
                            var hLink = parent.GetChildLink(item) as NomenclatureHierarchyLink;
                            xls[2, row].SetValue(hLink.Amount.ToString());
                        }
                        xls[2 + additionalCol, row].SetValue(itemUsage[1].ToString().Trim());
                        xls[3 + additionalCol, row].SetValue(itemUsage[0].ToString().Trim());
                        row++;
                    }
                }

                xls[1, 1, 3 + additionalCol, row - 1].Columns.AutoFit();
                xls[1, tableStart, 3 + additionalCol, row - tableStart].BorderTable();

                xls.SaveAs(fileName);
            }

            Enabled = true;
            Text = text;
            System.Diagnostics.Process.Start(fileName);

            DialogResult = DialogResult.OK;
        }

        private void searchButton_Click(object sender, EventArgs e)
        {
            searchedItemsListView.Items.Clear();

            if (searchTextBox.Text == string.Empty)
            {
                MessageBox.Show("Введите фрагмент Наименования для поиска");
                return;
            }

            // создаем ссылку на справочник "Номенклатура и изделия"
            ReferenceInfo refInfo = ReferenceCatalog.FindReference(NomenclatureReferenceGuid);
            Reference nomenclatureReference = refInfo.CreateReference();

            nomenclatureReference.LoadSettings.Clear();
            nomenclatureReference.LoadSettings.AddMasterGroupParameters();

            // находим тип "Прочее изделие"
            ClassObject otherItemClass = refInfo.Classes.Find(OtherItemClassGuid);
            // создаем фильтр поиска
            Filter filter = new Filter(refInfo);
            // добавляем условия поиска: "Тип = Прочее изделие" И "Наименование содержит [текст в поле searchTextBox]"
            filter.Terms.AddTerm(nomenclatureReference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, otherItemClass);
            filter.Terms.LogicalOperator = LogicalOperator.And;
            filter.Terms.AddTerm(nomenclatureReference.ParameterGroup[NameParameterGuid], ComparisonOperator.ContainsSubstring, searchTextBox.Text.Trim());
            // получаем список найденных объектов
            var items = nomenclatureReference.Find(filter).Select(t => new ListViewItem { Text = t.ToString(), Tag = t, Checked = true }).ToArray();

            itemsCount.Text = items.Count().ToString();
            itemsInList.Text = items.Count().ToString();
            searchedItemsListView.Items.AddRange(items);
        }

        private void searchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                searchButton_Click(this, EventArgs.Empty);
            }
        }

        private void excludeButton_Click(object sender, EventArgs e)
        {
            foreach (var item in searchedItemsListView.SelectedItems.Cast<ListViewItem>().ToArray())
                searchedItemsListView.Items.Remove(item);
            itemsInList.Text = searchedItemsListView.Items.Count.ToString();
        }

        private void selectAllButton_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in searchedItemsListView.Items)
                item.Checked = true;
        }

        private void unselectButton_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in searchedItemsListView.Items)
                item.Checked = false;
        }

        private void searchedItemsListView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                excludeButton_Click(sender, e);
            }
        }


    }
}
