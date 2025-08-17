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
using NomenclatureObject = TFlex.DOCs.Model.References.Nomenclature.NomenclatureObject;

using TFlex.Reporting;
using TFlex.DOCs.Model.Search;
using Button = System.Windows.Forms.Button;
using Filter = TFlex.DOCs.Model.Search.Filter;


namespace SummaryReport
{
    public partial class SelectionForm : Form
    {
        public SelectionForm()
        {
            InitializeComponent();
        }

        private void btnMakeReport_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        public void MakeReport()
        {

            var reportParameters = new ReportParameters();

            foreach (ListViewItem itmx in listViewObjectsReport.Items)
            {
                reportParameters.ListObjectCountId.Add(Convert.ToInt32(itmx.SubItems[0].Text.Trim()),
                                                       Convert.ToInt32(itmx.SubItems[3].Text.Trim()));
            }

            reportParameters.Assembly = chkAssembly.CheckState == CheckState.Checked ? true : false;
            reportParameters.Detail = chkDetail.CheckState == CheckState.Checked ? true : false;
            reportParameters.Complex = chkComplex.CheckState == CheckState.Checked ? true : false;
            reportParameters.Complement = chkComplement.CheckState == CheckState.Checked ? true : false;
            reportParameters.OtherItem = chkOther.CheckState == CheckState.Checked ? true : false;
            reportParameters.StandartItem = chkStandart.CheckState == CheckState.Checked ? true : false;
            reportParameters.Material = chkMaterial.CheckState == CheckState.Checked ? true : false;
            reportParameters.Documentation = chkDocumentation.CheckState == CheckState.Checked ? true : false;
            reportParameters.ComplexProgram = chkComplexProgram.CheckState == CheckState.Checked ? true : false;
            reportParameters.ComponentProgram = chkComponentProgram.CheckState == CheckState.Checked ? true : false;

            var report = new Report();

            report.MakeReport(context, reportParameters);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            TFlex.DOCs.UI.Objects.TFlexDOCsClientUI.Initialize();
            this.ClientSize = new System.Drawing.Size(546, 85);
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            this.btnMakeReport.Location = new System.Drawing.Point(12, 45);
            this.buttonClose.Location = new System.Drawing.Point(452, 45);
        }

        private Filter filter;



        private void FindObject(string denotation, ListView listView)
        {
            // соединяемся с БД T-FLEX DOCs 2012
            SqlConnection conn;
            var sqlConStringBuilder = new SqlConnectionStringBuilder();
            sqlConStringBuilder.DataSource = TFlex.DOCs.Model.ServerGateway.Connection.ServerName;
            sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
            sqlConStringBuilder.Password = "reportUser";
            sqlConStringBuilder.UserID = "reportUser";
            conn = new SqlConnection(sqlConStringBuilder.ToString());
            conn.Open();

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
                doc.Id = GetSqlInt32(reader, 0);
                doc.Naimenovanie = GetSqlString(reader, 1);
                doc.Oboznachenie = GetSqlString(reader, 2);
                objects.Add(doc);
            }
            reader.Close();

            var itmx = new ListViewItem[objects.Count];
            for (int i = 0; i < objects.Count; i++)
            {
                itmx[i] = new ListViewItem(objects[i].Id.ToString());

                itmx[i].SubItems.Add(objects[i].Naimenovanie);
                itmx[i].SubItems.Add(objects[i].Oboznachenie);
                itmx[i].SubItems.Add("1");
            }
            listView.Items.Clear();
            listView.Items.AddRange(itmx);

            #endregion
        }

        public string GetSqlString(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return String.Empty;

            return reader.GetString(field);
        }

        public int GetSqlInt32(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0;

            return reader.GetInt32(field);
        }

        public double GetSqlDouble(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0d;

            return reader.GetDouble(field);
        }

        private IReportGenerationContext context;


        private  Filter CreateFilter(ReferenceInfo documentReferenceInfo, List<string> listDenotation)
        {
            var filter = new Filter(documentReferenceInfo);
            filter.Terms.AddTerm("Обозначение", ComparisonOperator.IsOneOf, listDenotation);
            return filter;
        }

        private Filter CreateFilter(ReferenceInfo documentReferenceInfo, string denotation)
        {
            var filter = new Filter(documentReferenceInfo);
            filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
            return filter;
        }


        public void Init(IReportGenerationContext context)
        {
            this.context = context;
            var nomenclatureReference = new NomenclatureReference();
            var obj = (NomenclatureObject) nomenclatureReference.Find(context.ObjectsInfo[0].ObjectId);
            FindObject(obj.Denotation, listViewObjectsReport);
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
                    var denotations = new List<string>();

                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty; i++)
                    {
                        var filter = CreateFilter(nomenclatureReferenceInfo, xls[1, i].Text);
                        var reference = nomenclatureReferenceInfo.CreateReference();
                        var documentObjects = reference.Find(filter);
                        if (documentObjects == null)
                            MessageBox.Show("Объект " + xls[1, i].Value + " " + xls[2, i].Value + " не найден в базе");
                        else
                        {
                            var itmx = new ListViewItem[1];
                            itmx[0] = new ListViewItem(documentObjects[0].SystemFields.Id.ToString());

                            itmx[0].SubItems.Add(xls[1, i].Value);
                            itmx[0].SubItems.Add(xls[2, i].Value);
                            itmx[0].SubItems.Add(xls[3, i].Value + "");

                            listViewObjectsReport.Items.AddRange(itmx);
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

        private void loadButton_Click(object sender, EventArgs e)
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

                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty; i++)
                    {
                        var filter = CreateFilter(nomenclatureReferenceInfo, xls[1, i].Text);
                        var reference = nomenclatureReferenceInfo.CreateReference();
                        var documentObjects = reference.Find(filter);
                        if (documentObjects.Count == 0)
                            MessageBox.Show("Объект " + xls[1, i].Value + " " + xls[2, i].Value + " не найден в базе");
                        else
                        {
                            var itmx = new ListViewItem[1];
                            itmx[0] = new ListViewItem(documentObjects[0].SystemFields.Id.ToString());

                            itmx[0].SubItems.Add(xls[2, i].Value);
                            itmx[0].SubItems.Add(xls[1, i].Value);
                            itmx[0].SubItems.Add(xls[3, i].Value + "");

                            listViewObjectsReport.Items.AddRange(itmx);
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
                fileNameTextBox.Text = fileDialog.FileName;
            }
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            listViewObjectsReport.Items.Remove(listViewObjectsReport.FocusedItem);
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

     

        private void checkBoxForOneObject_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxForOneObject.CheckState == CheckState.Checked)
            {
                this.ClientSize = new System.Drawing.Size(546, 85);
                groupBox1.Visible = false;
                groupBox2.Visible = false;
                this.btnMakeReport.Location = new System.Drawing.Point(12, 45);
                this.buttonClose.Location = new System.Drawing.Point(452, 45);
            }
            else
            {
                this.ClientSize = new System.Drawing.Size(545, 600);
                groupBox1.Visible = true;
                groupBox2.Visible = true;
                this.btnMakeReport.Location = new System.Drawing.Point(12, 561);
                this.buttonClose.Location = new System.Drawing.Point(452, 561);
            }
        }

        private void chkAllTypes_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAllTypes.CheckState == CheckState.Checked)
            {
                chkAssembly.Checked = true;
                chkDetail.Checked = true;
                chkComplex.Checked = true;
                chkComplement.Checked = true;
                chkComplexProgram.Checked = true;
                chkComponentProgram.Checked = true;
                chkDocumentation.Checked = true;
                chkMaterial.Checked = true;
                chkOther.Checked = true;
                chkStandart.Checked = true;
            }
            else
            {
                chkAssembly.Checked = false;
                chkDetail.Checked = false;
                chkComplex.Checked = false;
                chkComplement.Checked = false;
                chkComplexProgram.Checked = false;
                chkComponentProgram.Checked = false;
                chkDocumentation.Checked = false;
                chkMaterial.Checked = false;
                chkOther.Checked = false;
                chkStandart.Checked = false;
            }
        }
    }
}

