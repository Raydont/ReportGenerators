using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Reflection;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.UI.Common;
using TFlex.DOCs.UI.Common.Filters;
using NomenclatureObject = TFlex.DOCs.Model.References.Nomenclature.NomenclatureObject;

using TFlex.Reporting;
using TFlex.DOCs.Model.Search;
using ParameterInfo = TFlex.DOCs.Model.Structure.ParameterInfo;


namespace TechnologicalDocumentationReport
{
    public partial class SelectionForm : Form
    {
        public SelectionForm()
        {
            InitializeComponent();

            var strings = Properties.Resources.TTP.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            //DataTable
            foreach (var line in strings)
            {
                var denotation = line.Substring(0, line.IndexOf(' '));
                var name = line.Substring(denotation.Length + 1, line.Length - denotation.Length - 1).Trim();
                gridViewTTP.Rows.Add(denotation, name);
            }

            strings = Properties.Resources.Complements.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            //DataTable
            foreach (var line in strings)
            {
                var denotation = line.Substring(0, line.IndexOf(' '));
                var name = line.Substring(denotation.Length + 1, line.Length - denotation.Length - 1).Trim();
                gridViewComplements.Rows.Add(denotation, name);
            }
        }


        private void btnMakeReport_Click(object sender, EventArgs e)
        {
            MakeReportParameters();
            DialogResult = DialogResult.OK;
        }


        public static string GetAutorLastName(IReportGenerationContext context)
        {
            PropertyInfo lastNameProperty = context.GetType().GetProperty("AuthorLastName");

            if (lastNameProperty != null)
            {
                return (string)lastNameProperty.GetValue(context, null);
            }

            PropertyInfo autorInfoProperty = context.GetType().GetProperty("AuthorInfo");
            if (autorInfoProperty != null)
            {
                var autorInfo = autorInfoProperty.GetValue(context, null);
                if (autorInfo != null)
                {
                    lastNameProperty = autorInfo.GetType().GetProperty("AuthorLastName");
                    if (lastNameProperty != null)
                    {
                        return (string)lastNameProperty.GetValue(autorInfo, null);
                    }
                }
            }

            return "";
        }

        string GetNameDocument(IReportGenerationContext context)
        {
            int documentID;

            // Выборка Id выделенного документа в T-FlexDocs 2010
            if (context.ObjectsInfo.Count == 1) documentID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return string.Empty;
            }
            // получение описания справочника          
            ReferenceInfo referenceInfo = ReferenceCatalog.FindReference(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
            // создание объекта для работы с данными
            Reference reference = referenceInfo.CreateReference();

            // получение объектов справочника          
            reference.Objects.Load();

            ReferenceObject changingObject = reference.Find(documentID);

            return changingObject.ParameterValues.GetParameter(new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")).ToString();
        }

        public void Init(IReportGenerationContext context)
        {
            SqlConnection connection = ApiDocs.GetConnection();

            textBoxAutor.Text = GetAutorLastName(context);

            //Запрос на выборку нормоконтроллёра 
            string allConstrQuery = String.Format
              (@"SELECT DISTINCT up2.ShortName
                 FROM ((UsersHierarchy uh INNER JOIN Users u ON (uh.s_ParentID = u.s_ObjectID)
                 INNER JOIN Users u2 ON (uh.s_ObjectID = u2.s_ObjectID)
                 INNER JOIN UserParameters up ON (up.s_ObjectID = u2.s_ObjectID)) INNER JOIN    
                 (UsersHierarchy uh2 INNER JOIN Users u3 ON (uh2.s_ParentID = u3.s_ObjectID)
                 INNER JOIN Users u4 ON (uh2.s_ObjectID = u4.s_ObjectID)) 
                 ON (u4.s_ObjectID = uh.s_ParentID)) INNER JOIN
                 ((UsersHierarchy uh3 INNER JOIN Users u5 ON (uh3.s_ParentID = u5.s_ObjectID)
                 INNER JOIN Users u6 ON (uh3.s_ObjectID = u6.s_ObjectID)) INNER JOIN
                 (UsersHierarchy uh4 INNER JOIN Users u7 ON (uh4.s_ParentID = u7.s_ObjectID)
                 INNER JOIN Users u8 ON (uh4.s_ObjectID = u8.s_ObjectID)
                 INNER JOIN UserParameters up2 ON (up2.s_ObjectID = u8.s_ObjectID))
                 ON(u6.s_ObjectID = uh4.s_ParentID))
                 ON (u3.s_ObjectID = uh3.s_ParentID)
                 WHERE up.LastName = N'{0}'
                 ORDER BY up2.ShortName", GetAutorLastName(context));

            SqlCommand allConstrCommand = new SqlCommand(allConstrQuery, connection);
            allConstrCommand.CommandTimeout = 0;
            SqlDataReader allConstrReader = allConstrCommand.ExecuteReader();


            if (allConstrReader.HasRows)
            {
                while (allConstrReader.Read())
                {
                    comboBoxNControl.Items.Add(allConstrReader.GetSqlString(0)); // combobox нормоконтроллёр
                    comboBoxProv1.Items.Add(allConstrReader.GetSqlString(0));
                    comboBoxProv2.Items.Add(allConstrReader.GetSqlString(0));

                }
            }
            comboBoxNControl.Text = comboBoxNControl.Items[0].ToString();
            comboBoxProv1.Text = comboBoxNControl.Items[0].ToString();
            comboBoxProv2.Text = comboBoxNControl.Items[0].ToString();
            allConstrReader.Close();
            textBoxNameDocument.Text = "Ведомость технологических документов изготовления\r\n".ToUpper() + GetNameDocument(context);

            var nomenclatureReference = new NomenclatureReference();
            var obj = (NomenclatureObject)nomenclatureReference.Find(context.ObjectsInfo[0].ObjectId);

            FindObject(obj.Denotation, listViewObjectsReport, buttonMoving);

            _context = context;

        }


        private void FindObject(string denotation, ListView listView, Button buttonMovObj)
        {
            // соединяемся с БД T-FLEX DOCs 2010
            var conn = ApiDocs.GetConnection();

            var objects = new List<TFDDocument>();

            #region Поиск объекта в номенклатуре

            string nomenclatureQuery = String.Format(@"SELECT   n.s_ObjectID,n.[Name],n.Denotation
                                                        FROM Nomenclature n
                                                        WHERE n.s_Deleted = 0 AND n.s_ActualVersion = 1 AND
                                                        n.Denotation = N'{0}'
                                                        ORDER BY n.[Name]",
                                                     denotation);
            var findObjectCommand = new SqlCommand(nomenclatureQuery, conn);
            findObjectCommand.CommandTimeout = 0;
            SqlDataReader reader = findObjectCommand.ExecuteReader();
            TFDDocument doc;

            while (reader.Read())
            {
                doc = new TFDDocument();
                doc.ObjectID = TechnologicalDocumentationReport.GetSqlInt32(reader, 0);
                doc.Denotation = TechnologicalDocumentationReport.GetSqlString(reader, 1);
                doc.Naimenovanie = TechnologicalDocumentationReport.GetSqlString(reader, 2);
                objects.Add(doc);
            }
            reader.Close();

            var itmx = new ListViewItem[objects.Count];
            for (int i = 0; i < objects.Count; i++)
            {
                itmx[i] = new ListViewItem(objects[i].ObjectID.ToString());

                itmx[i].SubItems.Add(objects[i].Naimenovanie);
                itmx[i].SubItems.Add(objects[i].Denotation);
            }
            listView.Items.Clear();
            listView.Items.AddRange(itmx);
            buttonMovObj.Enabled = false;

            #endregion
        }


        public ReportParameters reportParameters;

        private void MakeReportParameters()
        {
            reportParameters = new ReportParameters();

            reportParameters.AddLikeName.AddRange(nameList);
            reportParameters.AddLikeDenotation.AddRange(denotationList);


            reportParameters.listObjectId = new List<int>();
            reportParameters.NameDocument = textBoxNameDocument.Text.Trim();
            if (listViewObjectsReport.Items.Count > 0)
            {
                reportParameters.Denotation = listViewObjectsReport.Items[0].SubItems[1].Text;
            }
            else
            {
                reportParameters.Denotation = "";
            }
            reportParameters.DenotationDocument = textBoxDenotationDocument.Text.Trim();
            reportParameters.Author = textBoxAutor.Text.Trim();
            reportParameters.LiteraDocument = textBoxLitera.Text.Trim();
            reportParameters.NControl = comboBoxNControl.Text.Trim();
            reportParameters.Head = textBoxHead.Text.Trim();
            reportParameters.Checkman1 = comboBoxProv1.Text.Trim();
            reportParameters.Checkman2 = comboBoxProv2.Text.Trim();
            reportParameters.NameBoss = textBoxBoss.Text.Trim();
            if (checkBoxAddUpak.CheckState == CheckState.Checked)
                reportParameters.AddUpak = true;
            else
                reportParameters.AddUpak = false;


            foreach (ListViewItem itmx in listViewObjectsReport.Items)
            {
                reportParameters.listObjectId.Add(Convert.ToInt32(itmx.SubItems[0].Text.Trim()));
            }

            reportParameters.listDeleteObjectId = new List<int>();
            reportParameters.listDeleteObjectId.Add(0);

            foreach (ListViewItem itmx in listViewDelObj.Items)
            {
                reportParameters.listDeleteObjectId.Add(Convert.ToInt32(itmx.SubItems[0].Text.Trim()));
            }

            foreach (DataGridViewRow itmx in gridViewTTP.Rows)
            {
                var doc = new TFDDocument();
                doc.Denotation = (itmx.Cells["DenotationColumn"].Value + "").Trim();
                doc.Naimenovanie = (itmx.Cells["NameColumn"].Value + "").Trim();
                reportParameters.ListTTP.Add(doc);
            }

            foreach (DataGridViewRow itmx in gridViewComplements.Rows)
            {
                var doc = new TFDDocument();
                doc.Denotation = (itmx.Cells[0].Value + "").Trim();
                doc.Naimenovanie = (itmx.Cells[1].Value + "").Trim();
                reportParameters.ListComplement.Add(doc);
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }


        private void SelectionForm_Load(object sender, EventArgs e)
        {
            TFlex.DOCs.UI.Objects.TFlexDOCsClientUI.Initialize();
        }

        Filter filter;

        private void buttonFindObject_Click(object sender, EventArgs e)
        {
            //FindObject(textBoxFindObjeсt.Text, listViewFindObjects, buttonMoving);
            using (var filterDialog = new ReferenceFilterDialog())
            {
                var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
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
                    //Text = filter.ToString();
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
                    lSelectedObject.Text = "";
                }
            }
        }


        private void buttonMoving_Click(object sender, EventArgs e)
        {

            var itmx = new ListViewItem[1];

            itmx[0] = new ListViewItem(listViewFindObjects.FocusedItem.SubItems[0].Text);

            itmx[0].SubItems.Add(listViewFindObjects.FocusedItem.SubItems[1].Text);
            itmx[0].SubItems.Add(listViewFindObjects.FocusedItem.SubItems[2].Text);

            listViewObjectsReport.Items.AddRange(itmx);
        }

        private void listViewFindObjects_SelectedIndexChanged(object sender, EventArgs e)
        {
            var lv = sender as ListView;

            if ((listViewFindObjects.Items.Count != 0) && (listViewFindObjects.SelectedItems.Count != 0) && (listViewFindObjects.FocusedItem.Selected))
            {
                buttonMoving.Enabled = true;
                lSelectedObject.Text = lv.SelectedItems[0].SubItems[2].Text + " " + lv.SelectedItems[0].SubItems[1].Text;
            }
            else
            {
                buttonMoving.Enabled = false;
            }
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            listViewObjectsReport.Items.Remove(listViewObjectsReport.FocusedItem);
            buttonMoving.Enabled = false;
            lSelectedObject.Text = "";
        }

        private void buttonDeleteAll_Click(object sender, EventArgs e)
        {
            buttonMoving.Enabled = false;
            listViewObjectsReport.Items.Clear();
            lSelectedObject.Text = "";
        }



        private void buttonFindDelObj_Click(object sender, EventArgs e)
        {
            //FindObject(textBoxFindDeleteObj.Text, listViewFindDelObj, buttonMovingDelObj);
            using (var filterDialog = new ReferenceFilterDialog())
            {
                var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83"));
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
                    //Text = filter.ToString();
                    var nomenclatureReference = nomenclatureReferenceInfo.CreateReference();

                    var objects = nomenclatureReference.Find(filterDialog.Filter);
                    listViewFindDelObj.Items.Clear();

                    foreach (NomenclatureObject nomObject in objects)
                    {

                        var item = new ListViewItem(nomObject.SystemFields.Id.ToString());
                        item.SubItems.Add(nomObject.Name);
                        item.SubItems.Add(nomObject.Denotation);
                        item.Tag = nomObject;
                        listViewFindDelObj.Items.Add(item);
                    }
                    lDelSelectedObject.Text = "";
                }
            }
        }

        private void buttonMovingDelObj_Click(object sender, EventArgs e)
        {
            var itmx = new ListViewItem[1];

            itmx[0] = new ListViewItem(listViewFindDelObj.FocusedItem.SubItems[0].Text);

            itmx[0].SubItems.Add(listViewFindDelObj.FocusedItem.SubItems[1].Text);
            itmx[0].SubItems.Add(listViewFindDelObj.FocusedItem.SubItems[2].Text);

            listViewDelObj.Items.AddRange(itmx);
        }

        private void buttonDeleteDelObj_Click(object sender, EventArgs e)
        {
            listViewDelObj.Items.Remove(listViewDelObj.FocusedItem);
            buttonMovingDelObj.Enabled = false;
        }

        private void buttonDeleteAllDelObj_Click(object sender, EventArgs e)
        {
            buttonMovingDelObj.Enabled = false;
            listViewDelObj.Items.Clear();
        }

        private void listViewFindDelObj_SelectedIndexChanged(object sender, EventArgs e)
        {
            var lv = sender as ListView;
            if ((listViewFindDelObj.Items.Count != 0) && (listViewFindDelObj.SelectedItems.Count != 0) && (listViewFindDelObj.FocusedItem.Selected))
            {
                lDelSelectedObject.Text = lv.SelectedItems[0].SubItems[2].Text + " " + lv.SelectedItems[0].SubItems[1].Text;
                buttonMovingDelObj.Enabled = true;
            }
            else
            {
                buttonMovingDelObj.Enabled = false;
                lDelSelectedObject.Text = "";
            }
        }

        private void listViewObjectsReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            var lv = sender as ListView;

            if (lv.SelectedItems.Count > 0)
            {
                lSelectedObject.Text = lv.SelectedItems[0].SubItems[2].Text + " " + lv.SelectedItems[0].SubItems[1].Text;
            }
            else
            {
                lSelectedObject.Text = "";
            }
        }

        private void listViewDelObj_SelectedIndexChanged(object sender, EventArgs e)
        {
            var lv = sender as ListView;

            if (lv.SelectedItems.Count > 0)
            {
                lDelSelectedObject.Text = lv.SelectedItems[0].SubItems[2].Text + " " + lv.SelectedItems[0].SubItems[1].Text;
            }
            else
            {
                lDelSelectedObject.Text = "";
            }
        }

        private void AddConitionsNameCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (AddConitionsNameCheckBox.CheckState == CheckState.Checked)
            {
                nameGroupBox.Enabled = true;
                addFilterButton.Enabled = true;
            }
            else
            {
                nameGroupBox.Enabled = false;
                textBoxAddLikeName.Text = string.Empty;

                if (AddConitionsDenotationCheckBox.CheckState == CheckState.Unchecked)
                    addFilterButton.Enabled = false;
            }
        }

        private void AddConitionsDenotationCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (AddConitionsDenotationCheckBox.CheckState == CheckState.Checked)
            {
                denotationGroupBox.Enabled = true;
                addFilterButton.Enabled = true;
            }
            else
            {
                denotationGroupBox.Enabled = false;
                textBoxAddLikeDenotation.Text = string.Empty;

                if (AddConitionsNameCheckBox.CheckState == CheckState.Unchecked)
                    addFilterButton.Enabled = false;
            }
        }

        private List<string> nameList = new List<string>();
        private List<string> denotationList = new List<string>();
        private IReportGenerationContext _context;

        private void addFilterButton_Click(object sender, EventArgs e)
        {
            var itmx = new ListViewItem[1];

            if (AddConitionsNameCheckBox.CheckState == CheckState.Checked && AddConitionsDenotationCheckBox.CheckState == CheckState.Checked)
            {
                itmx[0] = new ListViewItem((filterListView.Items.Count + 1).ToString());

                itmx[0].SubItems.Add(textBoxAddLikeName.Text);
                itmx[0].SubItems.Add(textBoxAddLikeDenotation.Text);
                nameList.Add(textBoxAddLikeName.Text.Trim());
                denotationList.Add(textBoxAddLikeDenotation.Text.Trim());
                filterListView.Items.AddRange(itmx);

                return;
            }


            if (AddConitionsNameCheckBox.CheckState == CheckState.Checked)
            {
                itmx[0] = new ListViewItem((filterListView.Items.Count + 1).ToString());

                itmx[0].SubItems.Add(textBoxAddLikeName.Text);
                itmx[0].SubItems.Add("");

                denotationList.Add(null);
                nameList.Add(textBoxAddLikeName.Text.Trim());

                filterListView.Items.AddRange(itmx);
                return;
            }

            if (AddConitionsDenotationCheckBox.CheckState == CheckState.Checked)
            {
                itmx[0] = new ListViewItem((filterListView.Items.Count + 1).ToString());

                itmx[0].SubItems.Add("");
                itmx[0].SubItems.Add(textBoxAddLikeDenotation.Text);

                nameList.Add(null);
                denotationList.Add(textBoxAddLikeDenotation.Text.Trim());

                filterListView.Items.AddRange(itmx);
                return;
            }



        }

        private void tabPageAddParameters_Click(object sender, EventArgs e)
        {

        }

        private void listViewTTP_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
