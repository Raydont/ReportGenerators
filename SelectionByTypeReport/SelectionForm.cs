using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using ReportHelpers;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.UI.Common;
using TFlex.DOCs.UI.Common.Filters;
using NomenclatureObject = TFlex.DOCs.Model.References.Nomenclature.NomenclatureObject;

using TFlex.Reporting;
using TFlex.DOCs.Model.Search;
using Button = System.Windows.Forms.Button;
using Filter = TFlex.DOCs.Model.Search.Filter;
using TFlex.DOCs.Model.References;

namespace SelectionByTypeReport
{
    public partial class SelectionForm : Form
    {
        public SelectionForm()
        {
            InitializeComponent();
            selectionComboBox.SelectedIndex = 1;
        }

        private void chkAllObjects_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAllObjects.CheckState == CheckState.Checked)
            {
                chkPosition.Checked = true;
                chkZone.Checked = true;
                chkFormat.Checked = true;
                chkDenotation.Checked = true;
                chkName.Checked = true;
                chkCount.Checked = true;
                chkRemarks.Checked = true;
                chkId.Checked = true;
                chkLetter.Checked = true;
                chkCodeRCP.Checked = true;
                chkCreateDate.Checked = true;
                chkFirstUse.Checked = true;
                chkNomNumber.Checked = true;
                chKMeasureUnit.Checked = true;
                chkAuthorName.Checked = true;
                chkLastEditorName.Checked = true;
            }
            else
            {
                chkPosition.Checked = false;
                chkZone.Checked = false;
                chkFormat.Checked = false;
                chkDenotation.Checked = true;
                chkName.Checked = true;
                chkCount.Checked = true;
                chkRemarks.Checked = false;
                chkId.Checked = false;
                chkLetter.Checked = false;
                chkCodeRCP.Checked = false;
                chkCreateDate.Checked = false;
                chkFirstUse.Checked = false;
                chkNomNumber.Checked = false;
                chKMeasureUnit.Checked = false;
                chkAuthorName.Checked = false;
                chkLastEditorName.Checked = false;
            }

        }

        private void chkAllTypes_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAllTypes.CheckState == CheckState.Checked)
            {
                chkDocumentation.Checked = true;
                chkComplex.Checked = true;
                chkDetail.Checked = true;
                chkAssembly.Checked = true;
                chkStandart.Checked = true;
                chkOther.Checked = true;
                chkMaterial.Checked = true;
                chkComplement.Checked = true;
                chkComponentProgram.Checked = true;
                chkComplexProgram.Checked = true;
            }
            else
            {
                chkDocumentation.Checked = false;
                chkComplex.Checked = false;
                chkDetail.Checked = false;
                chkAssembly.Checked = false;
                chkStandart.Checked = false;
                chkOther.Checked = false;
                chkMaterial.Checked = false;
                chkComplement.Checked = false;
                chkComponentProgram.Checked = false;
                chkComplexProgram.Checked = false;

            }

        }


        private void btnMakeReport_Click(object sender, EventArgs e)
        {
            if ((chkDocumentation.CheckState == CheckState.Unchecked) && (chkComplex.CheckState == CheckState.Unchecked) && (chkDetail.CheckState == CheckState.Unchecked) &&
                (chkAssembly.CheckState == CheckState.Unchecked) && (chkStandart.CheckState == CheckState.Unchecked) && (chkOther.CheckState == CheckState.Unchecked) &&
                (chkMaterial.CheckState == CheckState.Unchecked) && (chkComplement.CheckState == CheckState.Unchecked) && (chkComplexProgram.CheckState == CheckState.Unchecked) &&
                (chkComponentProgram.CheckState == CheckState.Unchecked) && checkBoxOutputTechProcesses.CheckState == CheckState.Unchecked)
            {
                MessageBox.Show("Вы не выбрали ни одного типа объекта для отчета", "Внимание!!!", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }

            if ((chkName.CheckState == CheckState.Unchecked) && (chkDenotation.CheckState == CheckState.Unchecked) && (chkFormat.CheckState == CheckState.Unchecked) &&
               (chkZone.CheckState == CheckState.Unchecked) && (chkPosition.CheckState == CheckState.Unchecked) && (chkCount.CheckState == CheckState.Unchecked) &&
               (chkRemarks.CheckState == CheckState.Unchecked) && (chkLetter.CheckState == CheckState.Unchecked) && (chkId.CheckState == CheckState.Unchecked) &&
               (chkCodeRCP.CheckState == CheckState.Unchecked) && (chkFirstUse.CheckState == CheckState.Unchecked) && (chkCreateDate.CheckState == CheckState.Unchecked)
               && checkBoxOutputTechProcesses.CheckState == CheckState.Unchecked && chkNomNumber.CheckState == CheckState.Unchecked)
            {
                MessageBox.Show("Вы не выбрали ни одного поля для отчета", "Внимание!!!", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }

            DialogResult = DialogResult.OK;
        }

        public void MakeReport(Action<List<TFDDocument>, IndicationForm, LogDelegate> getDocuments = null)
        {

            var reportParameters = new ReportParameters();
            reportParameters.Position = (chkPosition.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Zone = (chkZone.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Format = (chkFormat.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Denotation = (chkDenotation.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Name = (chkName.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Count = (chkCount.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Remarks = (chkRemarks.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Letter = (chkLetter.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Id = (chkId.CheckState == CheckState.Checked) ? true : false;
            reportParameters.CodeRcp = (chkCodeRCP.CheckState == CheckState.Checked) ? true : false;
            reportParameters.FirstUse = (chkFirstUse.CheckState == CheckState.Checked) ? true : false;
            reportParameters.CreateDate = (chkCreateDate.CheckState == CheckState.Checked) ? true : false;
            reportParameters.NomenclatureNumber = (chkNomNumber.CheckState == CheckState.Checked) ? true : false;
            reportParameters.MeasureUnit = (chKMeasureUnit.CheckState == CheckState.Checked) ? true : false;
            reportParameters.AuthorName = (chkAuthorName.CheckState == CheckState.Checked) ? true : false;
            reportParameters.LastEditorName = (chkLastEditorName.CheckState == CheckState.Checked) ? true : false;

            reportParameters.Documentation = (chkDocumentation.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Complex = (chkComplex.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Detail = (chkDetail.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Assembly = (chkAssembly.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Standart = (chkStandart.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Other = (chkOther.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Materials = (chkMaterial.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Complement = (chkComplement.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ComplexProgram = (chkComplexProgram.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ComponentProgram = (chkComponentProgram.CheckState == CheckState.Checked) ? true : false;
            reportParameters.OutputTechProcesses = (checkBoxOutputTechProcesses.CheckState == CheckState.Checked) ? true : false;

            reportParameters.DetailTP = (chkBoxTPDetail.CheckState == CheckState.Checked) ? true : false;
            reportParameters.AssemblyTP = (chkBoxTPAssembly.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ComplementTP = (chkBoxTPKomplement.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ComplexTP = (chkBoxTPComplex.CheckState == CheckState.Checked) ? true : false;
            reportParameters.StandartItemsTP = (chkBoxTPStandartItems.CheckState == CheckState.Checked) ? true : false;

            reportParameters.WithOutKooperationDSE = (chkWithOutKooperationDSE.CheckState == CheckState.Checked) ? true : false;

            reportParameters.GroupByDevice = (radButGroupByDevices.Checked) ? true : false;
            reportParameters.GroupByTypes = (radButGroupByTypes.Checked) ? true : false;
            reportParameters.GroupByDeviceAndTypes = (radButGroupByDeviceAndTypes.Checked) ? true : false;

            reportParameters.AddLikeName.AddRange(nameList);
            reportParameters.AddLikeDenotation.AddRange(denotationList);
            reportParameters.listObjectCountId = new Dictionary<int, int>();

            reportParameters.LoadFolderName = _folderToSaveFiles;

            if (checkBoxOutputTechProcesses.CheckState == CheckState.Checked)
            {
                var denotations =
                    textBoxListObjectTP.Text.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Replace("\"", "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim())
                    .Distinct()
                    .ToList();

                var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);

                var filter = CreateFilter(nomenclatureReferenceInfo, denotations);
                var reference = nomenclatureReferenceInfo.CreateReference();
                var documentObjects = reference.Find(filter);

                foreach (var obj in documentObjects)
                {
                    reportParameters.listObjectCountId.Add(obj.SystemFields.Id, 1);
                }
            }
            else
            {
                foreach (ListViewItem itmx in listViewObjectsReport.Items)
                {
                    reportParameters.listObjectCountId.Add(Convert.ToInt32(itmx.SubItems[0].Text.Trim()), Convert.ToInt32(itmx.SubItems[3].Text.Trim()));
                }
            }

            if ((chkPosition.CheckState == CheckState.Unchecked) && (chkZone.CheckState == CheckState.Unchecked) && (chkFormat.CheckState == CheckState.Checked) &&
                (chkDenotation.CheckState == CheckState.Checked) && (chkName.CheckState == CheckState.Checked) && (chkCount.CheckState == CheckState.Checked) &&
                (chkRemarks.CheckState == CheckState.Checked) && (chkLetter.CheckState == CheckState.Unchecked) && (chkId.CheckState == CheckState.Unchecked) &&
                (chk5Col.CheckState == CheckState.Checked) && (chkCodeRCP.CheckState == CheckState.Unchecked) && (chkFirstUse.CheckState == CheckState.Unchecked) &&
                (chkNomNumber.CheckState == CheckState.Unchecked) && (chKMeasureUnit.CheckState == CheckState.Unchecked))
                reportParameters.Main5Col = true;
            else reportParameters.Main5Col = false;

            reportParameters.listDeleteObjectId = new List<int>();
            reportParameters.listDeleteObjectId.Add(0);

            foreach (ListViewItem itmx in listViewDelObj.Items)
            {
                reportParameters.listDeleteObjectId.Add(Convert.ToInt32(itmx.SubItems[0].Text.Trim()));
            }


            reportParameters.ExpansionNodes = (selectionComboBox.SelectedIndex == 1) ? true : false;

            SelectionByTypeReport.MakeReport(context, reportParameters, getDocuments);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            TFlex.DOCs.UI.Objects.TFlexDOCsClientUI.Initialize();
            if (ServerGateway.Connection.IsAdministrator)
            {
                btnMakeListAttachFiles.Visible = true;
                ExtractFilesByListButtin.Visible = true;
                buttonLoadFiles.Visible = true;
            }
            else
            {
                btnMakeListAttachFiles.Visible = false;
                ExtractFilesByListButtin.Visible = false;
                buttonLoadFiles.Visible = false;
            }
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

        private void FindObject(NomenclatureObject nomObj, ListView listView, Button buttonMovObj)
        {
            var doc = new TFDDocument(nomObj);
            var itmx = new ListViewItem[1];

            itmx[0] = new ListViewItem(doc.ObjectID.ToString());

            itmx[0].SubItems.Add(doc.Naimenovanie);
            itmx[0].SubItems.Add(doc.Denotation);
            itmx[0].SubItems.Add("1");

            listView.Items.Clear();
            listView.Items.AddRange(itmx);
            buttonMovObj.Enabled = false;

        }

        private void buttonMoving_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listViewObjectsReport.Items)
            {
                if (item.SubItems[0].Text.Trim() == listViewFindObjects.FocusedItem.SubItems[0].Text.Trim())
                {
                    item.SubItems[3].Text = (Convert.ToInt32(item.SubItems[3].Text.Trim()) + 1).ToString();
                    return;
                }
            }

            var itmx = new ListViewItem[1];

            itmx[0] = new ListViewItem(listViewFindObjects.FocusedItem.SubItems[0].Text);

            itmx[0].SubItems.Add(listViewFindObjects.FocusedItem.SubItems[1].Text);
            itmx[0].SubItems.Add(listViewFindObjects.FocusedItem.SubItems[2].Text);
            itmx[0].SubItems.Add("1");

            listViewObjectsReport.Items.AddRange(itmx);
        }

        private void listViewFindObjects_SelectedIndexChanged(object sender, EventArgs e)
        {
            var lv = sender as ListView;

            if ((listViewFindObjects.Items.Count != 0) && (listViewFindObjects.SelectedItems.Count != 0) && (listViewFindObjects.FocusedItem.Selected))
            {
                buttonMoving.Enabled = true;
                buttonMovingAll.Enabled = true;
                lSelectedObject.Text = lv.SelectedItems[0].SubItems[2].Text + " " + lv.SelectedItems[0].SubItems[1].Text;
            }
            else
            {
                buttonMoving.Enabled = false;
                buttonMovingAll.Enabled = false;
            }
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            listViewObjectsReport.Items.Remove(listViewObjectsReport.FocusedItem);
            buttonMoving.Enabled = false;
            buttonMovingAll.Enabled = false;
            lSelectedObject.Text = "";
        }

        private void buttonDeleteAll_Click(object sender, EventArgs e)
        {
            buttonMoving.Enabled = false;
            buttonMovingAll.Enabled = false;
            listViewObjectsReport.Items.Clear();
            lSelectedObject.Text = "";
        }

        private void chk4Col_CheckedChanged(object sender, EventArgs e)
        {
            if (chk5Col.CheckState == CheckState.Checked)
            {
                chkPosition.Checked = false;
                chkZone.Checked = false;
                chkFormat.Checked = true;
                chkDenotation.Checked = true;
                chkName.Checked = true;
                chkCount.Checked = true;
                chkRemarks.Checked = true;
                chkId.Checked = false;
                chkLetter.Checked = false;
                chkCodeRCP.Checked = false;
                chkFirstUse.Checked = false;
                chkCreateDate.Checked = false;
                chkNomNumber.Checked = false;
                chKMeasureUnit.Checked = false;
                chkAuthorName.Checked = false;
                chkLastEditorName.Checked = false;
            }
            else
            {
                chkPosition.Checked = false;
                chkZone.Checked = false;
                chkFormat.Checked = false;
                chkDenotation.Checked = true;
                chkName.Checked = true;
                chkCount.Checked = true;
                chkRemarks.Checked = false;
                chkId.Checked = false;
                chkLetter.Checked = false;
                chkCodeRCP.Checked = false;
                chkFirstUse.Checked = false;
                chkCreateDate.Checked = false;
                chkNomNumber.Checked = false;
                chKMeasureUnit.Checked = false;
                chkAuthorName.Checked = false;
                chkLastEditorName.Checked = false;
            }
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
        private IReportGenerationContext context;

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

        private string _folderToSaveFiles = "";

        private void btnMakeListAttachFiles_Click(object sender, EventArgs e)
        {
            var selectFolderDialog = new SelectFolderDialog();

            if (selectFolderDialog.ShowDialog(this) == DialogOpenResult.Ok)
            {
                _folderToSaveFiles = selectFolderDialog.SelectedPath;
                MakeReport(SaveFiles);
            }

        }

        private void SaveFiles(List<TFDDocument> documents, IndicationForm indicationForm, LogDelegate log)
        {
            var filesLink = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
            var fileNameGuid = new Guid("63aa0058-4a37-4754-8973-ffbc1b88f576");

            var documentReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Documents);

            var filer = CreateFilter(documentReferenceInfo, documents.Where(t => t.Level != 0).Select(t => t.Denotation).ToList());
            var reference = documentReferenceInfo.CreateReference();
            reference.LoadSettings.AddAllParameters();
            var filesRelarion = reference.LoadSettings.AddRelation(filesLink);
            filesRelarion.AddAllParameters();

            var documentObjects = reference.Find(filer);

            indicationForm.setStage(IndicationForm.Stages.DataProcessing);
            log("=== Получение и сохранение файлов ===");

            if (!_folderToSaveFiles.EndsWith("\\"))
            {
                _folderToSaveFiles = _folderToSaveFiles + "\\";
            }

            SelectionByTypeReport.HaveFiles = new List<string>();


            var loadFiles = MessageBox.Show("Выгружать файлы с итоговым отчетом(Да) или только создать итоговый отчет (Нет)", "", MessageBoxButtons.YesNo) == DialogResult.Yes;

            bool loadAllFiles = false;
            bool createFolders = false;
            bool loadOnlyLastPDF = false;

            if (loadFiles)
            {
                loadAllFiles =
                    MessageBox.Show("Выгрузить все файлы (Да) или только pdf-файлы (Нет)", "", MessageBoxButtons.YesNo) ==
                    DialogResult.Yes;

                if (!loadAllFiles)
                {
                    loadOnlyLastPDF = MessageBox.Show("Выгрузить только последний pdf?", "", MessageBoxButtons.YesNo) == DialogResult.Yes;

                }
                createFolders = MessageBox.Show("Создавать папки под каждый объект?", "", MessageBoxButtons.YesNo) == DialogResult.Yes;

            }

            foreach (var documentObject in documentObjects)
            {
                log("Объект: " + documentObject);

                var files = documentObject.GetObjects(filesLink);

                var denotation = documentObject[new Guid("b8992281-a2c3-42dc-81ac-884f252bd062")].GetString();

                var documentFolderName = _folderToSaveFiles + "\\" + (createFolders ? denotation + "\\" : "");


                var fileObjects = new List<FileReferenceObject>();
                if (loadAllFiles)
                    fileObjects.AddRange(files.OfType<FileReferenceObject>().ToList());
                else
                    fileObjects = files.OfType<FileReferenceObject>()
                    .Where(t => t[fileNameGuid].GetString().ToLower().EndsWith(".pdf"))
                    .ToList();

                if (loadOnlyLastPDF)
                {
                    fileObjects = fileObjects.OrderByDescending(t => t.SystemFields.CreationDate).Take(1).ToList();
                }


                if (fileObjects.Count > 0)
                {
                    SelectionByTypeReport.HaveFiles.Add(denotation);
                    if (createFolders)
                    {
                        if (!Directory.Exists(documentFolderName))
                        {
                            Directory.CreateDirectory(documentFolderName);
                        }
                    }
                    if (loadFiles)
                    {
                        foreach (var file in fileObjects)
                        {
                            var fileName = file[fileNameGuid].GetString();

                            log("Файл: " + fileName);

                            var filePath = documentFolderName + fileName;
                            file.GetHeadRevision();
                            File.Move(file.LocalPath, filePath);
                        }
                    }
                }
            }

            DialogResult = DialogResult.OK;
        }


        private void SaveFilesForRapidLoad(List<TFDDocument> documents, IndicationForm indicationForm, LogDelegate log)
        {
            var filesLink = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
            var fileNameGuid = new Guid("63aa0058-4a37-4754-8973-ffbc1b88f576");

            var documentReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);

            var filer = CreateFilter(documentReferenceInfo, documents/*.Where(t => t.Level != 0)*/.Select(t => t.Denotation).ToList());
            var reference = documentReferenceInfo.CreateReference();
            reference.LoadSettings.AddAllParameters();
         //   var filesRelarion = reference.LoadSettings.AddRelation(filesLink);
          //  filesRelarion.AddAllParameters();

            var documentObjects = reference.Find(filer);

            indicationForm.setStage(IndicationForm.Stages.DataProcessing);
            log("=== Получение и сохранение файлов ===");

            if (!_folderToSaveFiles.EndsWith("\\"))
            {
                _folderToSaveFiles = _folderToSaveFiles + "\\";
            }

            SelectionByTypeReport.HaveFiles = new List<string>();

            // var loadFiles = MessageBox.Show("Выгружать файлы (Да) или сформировать только Excel-файл со списком документов (Нет)", "", MessageBoxButtons.YesNo) ==
            //       DialogResult.Yes;
            var loadFiles = true; 
            bool createFolders = false;
            bool loadOnlyLastPDF = true;

            foreach (var documentObject in documentObjects)
            {
                log("Объект: " + documentObject);

                var files = new List<ReferenceObject>();

                if (((NomenclatureObject)documentObject).IsVersion)
                {
                    files.AddRange(((NomenclatureObject)documentObject).BaseVersion.LinkedObject.GetObjects(filesLink));
                }
                else
                {
                    files.AddRange(((NomenclatureObject)documentObject).LinkedObject.GetObjects(filesLink));
                }

                var denotation = ((NomenclatureObject)documentObject).LinkedObject[new Guid("b8992281-a2c3-42dc-81ac-884f252bd062")].GetString();
                var documentFolderName = _folderToSaveFiles + "\\" + (createFolders ? denotation + "\\" : "");

                var fileObjects = new List<FileReferenceObject>();

                fileObjects = files.OfType<FileReferenceObject>()
                .Where(t => t[fileNameGuid].GetString().ToLower().EndsWith(".pdf"))
                .ToList();

                if (loadOnlyLastPDF)
                {
                    fileObjects = fileObjects.OrderByDescending(t => t.SystemFields.CreationDate).Take(1).ToList();
                }


                if (fileObjects.Count > 0)
                {
                    SelectionByTypeReport.HaveFiles.Add(denotation);

                    if (loadFiles)
                    {
                        foreach (var file in fileObjects)
                        {
                            var fileName = file[fileNameGuid].GetString();
                            log("Файл: " + fileName);
                            var filePath = documentFolderName + fileName;
                            if (!File.Exists(filePath))
                            {
                                file.GetHeadRevision();
                                File.Move(file.LocalPath, filePath);
                            }
                        }
                    }
                }
            }
            DialogResult = DialogResult.Cancel;
        }

        private static Filter CreateFilter(ReferenceInfo documentReferenceInfo, List<string> listDenotation)
        {

            var filter = new Filter(documentReferenceInfo);
            filter.Terms.AddTerm("Обозначение", ComparisonOperator.IsOneOf, listDenotation);
            return filter;
        }

        private static Filter CreateFilter(ReferenceInfo documentReferenceInfo, string denotation)
        {

            var filter = new Filter(documentReferenceInfo);
            filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
            return filter;
            return filter;
        }


        public void Init(IReportGenerationContext context)
        {
            this.context = context;
            var nomenclatureReference = new NomenclatureReference(ServerGateway.Connection);
            var obj = (NomenclatureObject)nomenclatureReference.Find(context.ObjectsInfo[0].ObjectId);
            FindObject(obj, listViewObjectsReport, buttonMoving);
            textBoxListObjectTP.Text = obj.Denotation;
        }

        private void buttonMovingAll_Click(object sender, EventArgs e)
        {

            foreach (ListViewItem item in listViewObjectsReport.Items)
            {
                if (item.SubItems[0].Text.Trim() == listViewFindObjects.FocusedItem.SubItems[0].Text.Trim())
                {
                    item.SubItems[3].Text = (Convert.ToInt32(item.SubItems[3].Text) + 1).ToString();
                    return;
                }
            }

            foreach (ListViewItem item in listViewFindObjects.SelectedItems)
            {
                var itmx = new ListViewItem[1];

                itmx[0] = new ListViewItem(item.SubItems[0].Text);

                itmx[0].SubItems.Add(item.SubItems[1].Text);
                itmx[0].SubItems.Add(item.SubItems[2].Text);
                itmx[0].SubItems.Add("1");

                listViewObjectsReport.Items.AddRange(itmx);
            }


        }

        private class ReportDocument
        {
            public string Denotation { get; set; }

            public string Name { get; set; }

            public string Format { get; set; }
        }

        private void ExtractFilesByListButtin_Click(object sender, EventArgs e)
        {
            var dialog = new ExtractFilesByListDialog();
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _folderToSaveFiles = dialog.folderToSaveFiles;
                var indicationForm = new IndicationForm();
                indicationForm.TopMost = true;
                indicationForm.Visible = true;
                indicationForm.Activate();
                var log = new LogDelegate(indicationForm.writeToLog);

                indicationForm.setStage(IndicationForm.Stages.DataProcessing);
                log("=== Получение данных ===");

                var denotationList =
                    dialog.DenotationListTextBox.Text.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Replace("\"", "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim())
                    .Distinct()
                    .ToList();

                var filesLink = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");
                var fileNameGuid = new Guid("63aa0058-4a37-4754-8973-ffbc1b88f576");

                var documentReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);

                var filer = CreateFilter(documentReferenceInfo, denotationList);
                var reference = documentReferenceInfo.CreateReference();
                reference.LoadSettings.AddAllParameters();
             //   var filesRelarion = reference.LoadSettings.AddRelation(filesLink);
              //  filesRelarion.AddAllParameters();

                var documentObjects = reference.Find(filer);

                indicationForm.setStage(IndicationForm.Stages.DataProcessing);
                log("=== Получение и сохранение файлов ===");

                if (!_folderToSaveFiles.EndsWith("\\"))
                {
                    _folderToSaveFiles = _folderToSaveFiles + "\\";
                }

                var haveFiles = new List<string>();

                var createFolders = false;//MessageBox.Show("Создаввать папки под каждый объект?", "", MessageBoxButtons.YesNo) == DialogResult.Yes;


                foreach (var documentObject in documentObjects)
                {
                    log("Объект: " + documentObject);
                    var files = new List<ReferenceObject>();

                    if (((NomenclatureObject)documentObject).IsVersion)
                    {
                        files.AddRange(((NomenclatureObject)documentObject).BaseVersion.LinkedObject.GetObjects(filesLink));
                    }
                    else
                    {
                        files.AddRange(((NomenclatureObject)documentObject).LinkedObject.GetObjects(filesLink));
                    }

                    var denotation = ((NomenclatureObject)documentObject).LinkedObject[new Guid("b8992281-a2c3-42dc-81ac-884f252bd062")].GetString();

                    var documentFolderName = _folderToSaveFiles + "\\" + (createFolders ? denotation + "\\" : "");

                    var fileObject = files.OfType<FileReferenceObject>()
                        .Where(t => t[fileNameGuid].GetString().ToLower().EndsWith(".pdf")).OrderByDescending(t => t.SystemFields.CreationDate).FirstOrDefault();

                    if (fileObject != null)
                    {
                        haveFiles.Add(denotation.ToUpper());
                        if (createFolders)
                        {
                            if (!Directory.Exists(documentFolderName))
                            {
                                Directory.CreateDirectory(documentFolderName);
                            }
                        }

                        var fileName = fileObject[fileNameGuid].GetString();

                        log("Файл: " + fileName);

                        var filePath = documentFolderName + fileName;
                        fileObject.GetHeadRevision();

                        try
                        {
                            if (File.Exists(filePath))
                            {
                                File.Delete(filePath);
                            }

                            File.Move(fileObject.LocalPath, filePath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                    
                }

                Guid denotationGuid = new Guid("b8992281-a2c3-42dc-81ac-884f252bd062");
                Guid nameGuid = new Guid("7e115f38-f446-40ce-8301-9b211e6ce5fd");
                Guid formatGuid = new Guid("9781595e-7c37-42cb-8033-a9e192e7106b");

                var documntsList = documentObjects.Select(t => ((NomenclatureObject)t).LinkedObject).Select(t =>
                    new TFDDocument()
                    {
                        Denotation = t[denotationGuid].GetString().ToUpper(),
                        Naimenovanie = t[nameGuid].GetString(),
                        Format = t[formatGuid].GetString(),
                    })
                    .ToList();

                indicationForm.setStage(IndicationForm.Stages.ReportGenerating);
                log("=== Формирование отчета ===");

                var xls = new Xls();
                var reportFile = _folderToSaveFiles + "!отчет.xls";

                int row = 1;
                xls[1, row].SetValue("Формат");
                xls[2, row].SetValue("Обозначение");
                xls[3, row].SetValue("Наименование");
                xls[4, row].SetValue("Объект найден");
                xls[5, row].SetValue("Файл сохранен");
                row++;

                foreach (var denotation in denotationList)
                {
                    log("Вывод объекта: " + denotation);
                    var document = documntsList.FirstOrDefault(t => t.Denotation.ToUpper() == denotation.ToUpper());
                    xls[1, row].SetValue(document == null ? "-" : document.Format);
                    xls[2, row].SetValue(denotation);
                    xls[3, row].SetValue(document == null ? "-" : document.Naimenovanie);
                    xls[4, row].SetValue(document == null ? "-" : "+");
                    xls[5, row].SetValue(haveFiles.Any(t => t.ToUpper() == denotation.ToUpper()) ? "+" : "-");
                    row++;
                }

                xls[1, 1, 5, 1].Font.Bold = true;
                xls[1, 1, 5, row - 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                xls.AutoWidth();

                var reportParameters = new ReportParameters();
                reportParameters.Name = true;
                reportParameters.Denotation = true;
                reportParameters.Format = true;

                try
                {
                    xls.SaveAs(reportFile);
                }
                catch
                { }

                xls.Visible = true;

                indicationForm.Dispose();

            }
        }

        private void chkBoxTPAllTypes_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBoxTPAllTypes.CheckState == CheckState.Checked)
            {
                chkBoxTPComplex.Checked = true;
                chkBoxTPDetail.Checked = true;
                chkBoxTPAssembly.Checked = true;
                chkBoxTPStandartItems.Checked = true;
                chkBoxTPKomplement.Checked = true;
            }
            else
            {
                chkBoxTPComplex.Checked = false;
                chkBoxTPDetail.Checked = false;
                chkBoxTPAssembly.Checked = false;
                chkBoxTPStandartItems.Checked = false;
                chkBoxTPKomplement.Checked = false;
            }
        }

        private void chkBoxTPDetail_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxOutputTechProcesses_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxOutputTechProcesses.CheckState == CheckState.Checked)
            {
                groupBox4.Enabled = true;
            }
            else
            {
                groupBox4.Enabled = false;
            }
        }

        private void listViewObjectsReport_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var countDeviceForm = new CountDeviceForm();

            countDeviceForm.numericUpDownCountDevice.Value = Convert.ToInt32(listViewObjectsReport.SelectedItems[0].SubItems[3].Text);
            countDeviceForm.groupBox1.Text = listViewObjectsReport.SelectedItems[0].SubItems[2].Text + " " + listViewObjectsReport.SelectedItems[0].SubItems[1].Text;

            if (countDeviceForm.ShowDialog(this) == DialogResult.OK)
            {
                listViewObjectsReport.SelectedItems[0].SubItems[3].Text = countDeviceForm.numericUpDownCountDevice.Value.ToString();
            }

        }

        private void loadExcelButton_Click(object sender, EventArgs e)
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
                    var badObjects = new List<string>();
                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty; i++)
                    {
                        var denotation = xls[1, i].Text.Trim();
                        var filter = CreateFilter(nomenclatureReferenceInfo, denotation);
                        var itemInList = listViewObjectsReport.Items.OfType<ListViewItem>().FirstOrDefault(t => t.SubItems[1].Text == denotation);
                        if (itemInList != null)
                        {
                            var count = Convert.ToInt32(itemInList.SubItems[3].Text) + Convert.ToInt32(xls[3, i].Value + "");
                            itemInList.SubItems[3].Text = count + "";
                            continue;
                        }

                        var documentObjects = reference.Find(filter);
                        if (documentObjects == null || documentObjects.Count == 0)
                        {
                            badObjects.Add(xls[1, i].Value + " " + xls[2, i].Value);
                           
                        }
                        else
                        {
                            var itmx = new ListViewItem[1];
                            itmx[0] = new ListViewItem(documentObjects[0].SystemFields.Id.ToString());


                            {
                                itmx[0].SubItems.Add(denotation);
                                itmx[0].SubItems.Add(xls[2, i].Value + "");
                                itmx[0].SubItems.Add(xls[3, i].Value + "");
                            }

                            listViewObjectsReport.Items.AddRange(itmx);
                        }
                    }
                    if (badObjects.Count > 0)
                    {
                        MessageBox.Show("Объекты не найденные в Номенклатуре: " + string.Join("\r\n", badObjects));
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



        //        List<string> listDenotation = new List<string>
        //        {
        //"БФМИ.468381.004",
        //"БФМИ.468171.041",
        //"БФМИ.467871.017",
        //"БФМИ.468214.096",
        //"БФМИ.468173.094",
        //"БФМИ.468151.375",
        //"БФМИ.468353.135",
        //"БФМИ.469172.071",
        //"БФМИ.685619.064-04",
        //"БФМИ.685619.195",
        //"БФМИ.685619.198",
        //"БФМИ.685619.196",
        //"БФМИ.685619.255",
        //"БФМИ.685619.197",
        //"БФМИ.685619.224",
        //"БФМИ.685619.223",
        //"БФМИ.685619.230",
        //"БФМИ.685619.231",
        //"БФМИ.685619.232",
        //"БФМИ.685613.278",
        //"БФМИ.685619.233",
        //"БФМИ.685619.234",
        //"БФМИ.685619.235",
        //"БФМИ.685619.229",
        //"БФМИ.685619.236",
        //"БФМИ.685619.236-01",
        //"БФМИ.685619.228",
        //"БФМИ.685613.279",
        //"БФМИ.685619.227",
        //"БФМИ.685619.208",
        //"БФМИ.685613.209-09",
        //"БФМИ.685614.146",
        //"БФМИ.685614.147",
        //"БФМИ.685619.212",
        //"БФМИ.685664.016",
        //"БФМИ.685619.207",
        //"БФМИ.685613.283",
        //"БФМИ.685613.286",
        //"БФМИ.685619.215",
        //"БФМИ.685619.210",
        //"БФМИ.685619.225",
        //"БФМИ.685619.225-01",
        //"БФМИ.685619.240",
        //"БФМИ.685619.241",
        //"БФМИ.685619.242",
        //"БФМИ.685619.237",
        //"БФМИ.685619.241-01",
        //"БФМИ.685612.521",
        //"БФМИ.685611.948-03",
        //"БФМИ.685614.219",
        //"БФМИ.685619.213",
        //"БФМИ.685619.214",
        //"БФМИ.685664.015",
        //"БФМИ.685661.156",
        //"БФМИ.685613.271",
        //"БФМИ.685613.272",
        //"БФМИ.685613.273",
        //"БФМИ.685616.015",
        //"БФМИ.685612.518",
        //"БФМИ.685613.274",
        //"БФМИ.685613.275",
        //"БФМИ.685615.050",
        //"БФМИ.685619.216",
        //"БФМИ.685613.284",
        //"БФМИ.685619.217",
        //"БФМИ.685619.218",
        //"БФМИ.685619.209",
        //"БФМИ.685613.270",
        //"БФМИ.685619.219",
        //"БФМИ.685619.220",
        //"БФМИ.685664.014",
        //"БФМИ.685619.226",
        //"БФМИ.685619.238",
        //"БФМИ.685619.221",
        //"БФМИ.685613.276",
        //"БФМИ.685613.277",
        //"БФМИ.685616.016",
        //"БФМИ.685614.144",
        //"БФМИ.685619.239",
        //"БФМИ.685619.211",
        //"БФМИ.685614.145"
        //};


        private void buttonLoadFiles_Click(object sender, EventArgs e)
        {
            ВыгрузкаДляКооперации();
        }

        private void ВыгрузкаДляКооперации()
        {
            var dialog = new ExtractFilesByListDialog();
            List<string> listDenObjects = new List<string>();

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _folderToSaveFiles = dialog.folderToSaveFiles + @"\Выгрузка " + DateTime.Now.ToShortDateString() + @"\";

                listDenObjects = dialog.DenotationListTextBox.Text.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Replace("\"", "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim())
                    .Distinct()
                    .ToList();
            }

            var beginFolder = _folderToSaveFiles;
            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);

            listViewObjectsReport.Items.Clear();
            var badObjects = new List<string>();

            foreach (var denotation in listDenObjects)
            {
                var reference = nomenclatureReferenceInfo.CreateReference();
                var filter = CreateFilter(nomenclatureReferenceInfo, denotation);

                var documentObject = reference.Find(filter).FirstOrDefault();
                if (documentObject == null)
                {
                    badObjects.Add(denotation);
                }

                else
                {
                    var itmx = new ListViewItem[1];
                    itmx[0] = new ListViewItem(documentObject.SystemFields.Id.ToString());
                    itmx[0].SubItems.Add("");
                    itmx[0].SubItems.Add(denotation);
                    itmx[0].SubItems.Add("1");

                    listViewObjectsReport.Items.AddRange(itmx);
                    var nameDocumentObject = documentObject.ToString().Length > 100 ? documentObject[Guids.Nomenclature.Fields.Denotation].ToString() : documentObject.ToString();
                    _folderToSaveFiles = beginFolder + nameDocumentObject + @"\";
                    Directory.CreateDirectory(_folderToSaveFiles);
                    MakeReport(SaveFilesForRapidLoad);
                    MakeReport();
                    listViewObjectsReport.Items.Clear();
                }
            }
            if (badObjects.Count > 0)
            {
                Clipboard.SetText(string.Join("\r\n", badObjects));
                MessageBox.Show("Объекты, не найденные в номенклатуре\r\n" + string.Join("\r\n", badObjects));
            }
            DialogResult = DialogResult.Cancel;
        }

        private void ВыгрузкаДокументовЭКД()
        {
            var dialog = new ExtractFilesByListDialog();

            List<string> listDenObjects = new List<string>();

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                _folderToSaveFiles = dialog.folderToSaveFiles + @"\Выгрузка " + DateTime.Now.ToShortDateString() + @"\";

                listDenObjects = Clipboard.GetText().Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(t => t.Replace("\"", "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim())
                    .Distinct()
                    .ToList();
                /*  dialog.DenotationListTextBox.Text.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)
                  .Select(t => t.Replace("\"", "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim())
                  .Distinct()
                  .ToList();*/
            }

            var beginFolder = _folderToSaveFiles;

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);

            listViewObjectsReport.Items.Clear();
            var badObjects = new List<string>();
            var reportObjects = new List<ReferenceObject>();

            foreach (var denotation in listDenObjects)
            {
                var reference = nomenclatureReferenceInfo.CreateReference();

                var filter = CreateFilter(nomenclatureReferenceInfo, denotation);

                var documentObject = reference.Find(filter).FirstOrDefault();
                if (documentObject == null)
                {
                    badObjects.Add(denotation);
                }
                else
                {
                    reportObjects.Add(documentObject);
                }
            }

            if (badObjects.Count > 0)
            {
                Clipboard.SetText(string.Join("\r\n", badObjects));
                MessageBox.Show("Объекты, не найденные в номенклатуре\r\n" + string.Join("\r\n", badObjects));
            }

            foreach (var reportObject in reportObjects)
            {
                var itmx = new ListViewItem[1];
                itmx[0] = new ListViewItem(reportObject.SystemFields.Id.ToString());
                itmx[0].SubItems.Add("");
                itmx[0].SubItems.Add(reportObject[Guids.Nomenclature.Fields.Denotation].GetString());

                itmx[0].SubItems.Add("1");

                listViewObjectsReport.Items.AddRange(itmx);
            }
            _folderToSaveFiles = beginFolder + "Выгрузка документов ЭКД" + @"\";
            Directory.CreateDirectory(_folderToSaveFiles);

             MakeReport(SaveFilesForRapidLoad);
             MakeReport();
             listViewObjectsReport.Items.Clear();
           
            DialogResult = DialogResult.Cancel;
        }

        private void chkWithOutKooperationDSE_CheckedChanged(object sender, EventArgs e)
        {
            if (selectionComboBox.Text.Trim() == "Без разузловки")
            {
                chkWithOutKooperationDSE.Checked = false;
            }
        }

        private void selectionComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (selectionComboBox.Text.Trim() == "Без разузловки")
            {
                chkWithOutKooperationDSE.Checked = false;
            }
        }
    }
}

