using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using TFlex.DOCs.Model.References.Files;
using System.Text.RegularExpressions;

namespace ListElementsReport
{
    public partial class SelectionForm : Form
    {


        // Гуид справочника "Группы и пользователи"
        private static readonly Guid UsersAndGroupsReferenceGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        //Тип справочника "Группы и пользователи" должность
        private static readonly Guid TypePostUsersGuid = new Guid("ea7c8581-c329-48bc-9e54-9becdccc5ae7");
        //параметр Наименование объекта справочника "Группыи пользователи"
        private static readonly Guid NameParametrGuid = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
        //параметр Короткое имя справочника "Группы и пользователи"
        private static readonly Guid LastNameGuid = new Guid("76a97c36-f2d6-49ad-abb0-f5fdc91c071b");

        private IReportGenerationContext _context;
        public ReportParameters reportParameters;
        public SelectionForm()
        {
            InitializeComponent();
        }

        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        private void buttonLoadFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxFileName.Text = openFileDialog.FileName;
                if (!textBoxFileName.Text.ToLower().Contains(".atr") && !textBoxFileName.Text.ToLower().Contains(".xls"))
                    return;
            }
            else
                return;

            connectDB();

            if ((textBoxFileName.Text.ToLower().Contains(".atr") && radioButtonPCAD.Checked) || 
                (textBoxFileName.Text.ToLower().Contains(".xls") && radioButtonAltiumDesigner.Checked))
            {
                makeReportButton.Enabled = true;
            }
               
        }

        private void makeReportButton_Click(object sender, EventArgs e)
        {
            if (MakeReportParameters())
                DialogResult = DialogResult.OK;
        }

        private bool MakeReportParameters()
        {
            reportParameters = new ReportParameters();
          
            reportParameters.DenotationDevice = textBoxDenotationDevice.Text.Trim();
            reportParameters.FirstUse = textBoxFirstUse.Text.Trim();
            reportParameters.filePath = openFileDialog.FileName;
            if (openFileDialog.FileName.Contains(".atr"))
            {
                try
                {
                    var data = File.ReadAllLines(openFileDialog.FileName, Encoding.GetEncoding("Windows-1251")).ToList();
                    if (data.Count > 0 && !data[0].Contains(";"))
                    {
                        MessageBox.Show("Некорректные разделитель в файле. Корректным разделителем является \";\".", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    for (var i = data.Count - 1; i > 0; i--)
                    {
                        if (!data[i].StartsWith("\""))
                        {
                            data[i - 1] += "\r\n" + data[i];
                            data.RemoveAt(i);
                        }
                    }

                    var numberColumns = new NumbersColumns(data[0]);
                    if (numberColumns.Designator == -1 || numberColumns.Type == -1 || numberColumns.Value == -1)
                    {
                        MessageBox.Show(data[0] + "Отчет не может быть сформирован, поскольку не хватает данных из файла ", "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    for (int i = 1; i < data.Count; i++)
                    {
                       
                        try
                        {
                            //"RefDes"; "Type"; "Value"; "Допуск"; "TU"; "Обозначение"; "Диап. номин."
                            //"C1"; "K53-68"D"-16B"; "47мк"; "±20%"; "АЖЯР.673546.007ТУ"; ""; "22 33 47 68 100мкФ"

                            var parameters = data[i].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                            if (parameters.Length == 0)
                            {
                                continue;
                            }

                            reportParameters.FileData.Add(new FileData(parameters, numberColumns));
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Обратитесь в отдел 911. Перечень элементов не может быть сформирован!\n"
                                  + "Некорректные данные. Ошибка в строке № " + i + "\n" + data[i],
                                  "Ошибка в исходном файле!\r\n" + ex, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show("Выберите корректный файл!\n" + e.ToString(), "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            else
            {
                var errorId = 0;
                var xls = new ListElementReport.Xls();
                try
                {
                    if (openFileDialog.FileName.Contains(".xls"))
                    {
                        errorId = 1;
                        xls.Application.DisplayAlerts = false;
                        xls.Open(openFileDialog.FileName, false);
                        //Переход на 1 лист
                        xls.SelectWorksheet(1);

                        var numberColumns = new NumbersColumns(xls);

                        if (numberColumns.Type == -1 && numberColumns.Designator == -1)
                        {
                            MessageBox.Show("Отчет не может быть сформирован, поскольку не хватает данных из файла", "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }

                        for (int i = 2; xls[numberColumns.Type, i].Text.Trim() != string.Empty; i++)
                        {
                            reportParameters.FileData.Add(new FileData(xls, numberColumns, i));
                        }                    
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(errorId + " Выберите корректный файл!\n" + e.ToString(), "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                finally
                {
                    xls.Quit(true);

                }
            }

            reportParameters.AuthorBy = tbDesignedBy.Text.Trim();
            reportParameters.MusteredBy = cbControlledBy.Text.Trim();
            reportParameters.NControlBy = tbNControlBy.Text.Trim();
            reportParameters.ApprovedBy = tbApprovedBy.Text.Trim();
            reportParameters.CountHeadRows = Convert.ToInt32(nudCountHeadRows.Value);
            reportParameters.BeginPage = Convert.ToInt32(nudPageBeginReport.Value);


            if (cbLetter.Text.Trim() != "-" && cbLetter.Text.Length > 0)
                reportParameters.Litera1 = cbLetter.Text[0];

            if (cbLetter.Text.Trim() != "-" && cbLetter.Text.Length > 1)
                reportParameters.Litera2 = cbLetter.Text[1];

            reportParameters.NumberSol = tbNumberSol.Text.Trim();
            reportParameters.Code = tbCode.Text.Trim();
            reportParameters.NameDevice = textBoxName.Text.Trim();

            reportParameters.AuthorCheck = (checkBoxAuthor.CheckState == CheckState.Checked) ? true : false;
            reportParameters.MusteredBCheck = (checkBoxMasteredBy.CheckState == CheckState.Checked) ? true : false;
            reportParameters.NControlByCheck = (checkBoxNControllBy.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ApprovedByCheck = (checkBoxApprovedBy.CheckState == CheckState.Checked) ? true : false;
            reportParameters.SelectTKS = (checkBoxSelectionTKS.CheckState == CheckState.Checked) ? true : false;
            reportParameters.IsLoadKre = (checkBoxLoadKre.CheckState == CheckState.Checked) ? true : false;
            reportParameters.IsManyDevice = (checkBoxManyDevice.CheckState == CheckState.Checked) ? true : false;
            reportParameters.IsKonstr = (checkBoxKonstr.CheckState == CheckState.Checked) ? true : false;

            reportParameters.AuthorByDate = dTPAuthor.Value;
            reportParameters.MusteredByDate = dTPMasteredBy.Value;
            reportParameters.NControlByDate = dTPNControlBy.Value;
            reportParameters.ApprovedByDate = dTPApprovedBy.Value;

            reportParameters.dBElements.AddRange(dBElements);

            if (reportParameters.FileData.Any(t => !string.IsNullOrEmpty(t.Наименование)) && reportParameters.FileData.Any(t => ! string.IsNullOrEmpty(t.Обозначение)) && !reportParameters.IsManyDevice)
            {
                if (MessageBox.Show("Флаг СВЧ не отмечен, данные по Обозначению и Наименованию считаны. Вывести ПЭ3 в режиме СВЧ?", "Внимание!", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) ==
                    DialogResult.OK)
                {
                    reportParameters.IsManyDevice = true;
                }
            }

            if (reportParameters.FileData.All(t => string.IsNullOrEmpty(t.Наименование)) && reportParameters.FileData.All(t => string.IsNullOrEmpty(t.Обозначение)) && reportParameters.IsManyDevice)
            {
                MessageBox.Show("Флаг СВЧ отмечен, но данные по Обозначению и Наименованию отсутствуют. ПЭ3 будет сформировано без режима СВЧ.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                reportParameters.IsManyDevice = false;
            }

            return true;

        }
      

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            //Создаем ссылку на справочник Группы и пользователи
            ReferenceInfo info = ReferenceCatalog.FindReference(UsersAndGroupsReferenceGuid);
            Reference reference = info.CreateReference();

            //Находим тип «Должность»
            ClassObject classObject = info.Classes.Find(TypePostUsersGuid);
            //Создаем фильтр
            Filter filterPost = new Filter(info);
            //Добавляем условие поиска – «Тип = Должность»
            filterPost.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.Equal, classObject);

            //Получение списка конструкторов из своего подразделения
            var myParentDesign = ClientView.Current.GetUser().Parents.Where(t => t[NameParametrGuid].Value.ToString() == "Конструктор").ToList();
            if (myParentDesign.Count() == 1)
            {
                var listMyDepartament = myParentDesign[0].Children.OrderBy(t => t[LastNameGuid].Value).Select(t => t[LastNameGuid].Value.ToString()).ToArray();
                cbControlledBy.Items.AddRange(listMyDepartament);
                cbControlledBy.Text = listMyDepartament[0];
            }

            tbDesignedBy.Text = ClientView.Current.GetUser()[LastNameGuid].Value.ToString();

            //var myParents = ClientView.Current.GetUser().Parents;

            //foreach (var parent in myParents)
            //{
            //    foreach (var parentParent in parent.Parents)
            //    {

            //        foreach (var child in parentParent.Children)
            //        {
            //            if (child[NameParametrGuid].Value.ToString() == "Начальник")
            //            {
            //                tbApprovedBy.Text = child.Children[0][LastNameGuid].Value.ToString();
            //            }
            //        }
            //    }
            //}
        }
        private List<DataBaseElements> dBElements = new List<DataBaseElements>();

        private void buttonLoadDB_Click(object sender, EventArgs e)
        {
        }
        private void connectDB()
        {
            dBElements = new List<DataBaseElements>();

            FileReference reference = new FileReference(ServerGateway.Connection);
            var fileObject = (FileReferenceObject)reference.FindByPath(@"Служебные файлы\Перечень элементов\base.MDB");
            fileObject.GetHeadRevision();

            //    File.Copy(fileObject.LocalPath, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\localbase.MDB", true);
            // OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + @"\\192.168.5.7\DBRB218\base.MDB" + "\"");
            Clipboard.SetText(fileObject.LocalPath);
            using (OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"" + fileObject.LocalPath + "\""))
            // using (OleDbConnection con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + fileObject.LocalPath + "\""))
            {
                con.Open();

                var tables = con.GetSchema("Tables");

                var listTables = new List<string>();
                foreach (DataRow tableInfo in tables.Rows)
                {
                    var tableName = tableInfo[tables.Columns["TABLE_NAME"]] + "";
                    var tableColumns = con.GetSchema("Columns", new[] { null, null, tableName });
                    var tableColumnsList = new List<string>();
                    foreach (DataRow tableColumnInfo in tableColumns.Rows)
                    {
                        var columnName = tableColumnInfo[tableColumns.Columns["COLUMN_NAME"]] + "";
                        tableColumnsList.Add(columnName);
                    }

                    if (tableColumnsList.Any(t => t.ToLower().Trim() == "comment") &&
                      tableColumnsList.Any(t => t.ToLower().Trim() == "description") &&
                      tableColumnsList.Any(t => t.ToLower().Trim() == "part number") &&
                      tableColumnsList.Any(t => t.ToLower().Trim() == "tu") &&
                      tableColumnsList.Any(t => t.ToLower().Trim() == "range") &&
                      tableColumnsList.Any(t => t.ToLower().Trim() == "optional parameter") &&
                      tableColumnsList.Any(t => t.ToLower().Trim() == "tolerance") &&
                      tableColumnsList.Any(t => t.ToLower().Trim() == "value"))
                    //|| tableColumnsList.Any(t => t.ToLower().Trim() == "ту")))

                    //  tableColumnsList.Any(t => t.ToLower().Trim() == "допуск"))//&&
                    //  (tableColumnsList.Any(t => t.ToLower().Trim() == "диап_номин_")|| tableColumnsList.Any(t => t.ToLower().Trim() == "диап_ номин_")))
                    {
                        listTables.Add(tableName);
                    }
                }
                //  dataGridView1.DataSource = tables;
                int countLinkedObj = 0;
                foreach (var nameTable in listTables)
                {
                    OleDbCommand cmd = con.CreateCommand();
                    OleDbDataReader dataReader;
                    try
                    {
                        var query = string.Format(@"SELECT mainTable.Comment, 
                                                   mainTable.Description, 
                                                   mainTable.[Part Number], 
                                                   mainTable.TU, 
                                                   mainTable.Range,  
                                                   mainTable.[Optional parameter],
                                                   mainTable.Tolerance,
                                                   mainTable.Value,
                                                   link.[Код детали],
                                                   link.[Количество]
                                                   FROM [{0}] AS mainTable LEFT JOIN [{1}] 
                                                   AS link ON mainTable.[Part Number] = link.[Код ПКИ]",
                                                           nameTable, "Связь " + nameTable + " к Связанным объектам");

                        cmd.CommandText = query;
                        cmd.Connection = con;
                        dataReader = cmd.ExecuteReader();
                        // MessageBox.Show(query);
                        //Clipboard.SetText(query);
                    }
                    catch
                    {

                        var query = string.Format(@"SELECT Comment, 
                                                   Description, 
                                                   [Part Number], 
                                                   TU, 
                                                   Range,  
                                                   [Optional parameter],
                                                   Tolerance,
                                                   Value,
                                                   ''
                                                   FROM [{0}]", nameTable);
                        cmd.CommandText = query;
                        cmd.Connection = con;
                        dataReader = cmd.ExecuteReader();



                    }
                    var prewElement = string.Empty;
                    var dBElement = new DataBaseElements();

                    while (dataReader.Read())
                    {
                        var linkedObject = new LinkedObjectDB();
                        if (dataReader[2].ToString() != prewElement)
                        {
                            dBElement = new DataBaseElements();
                            dBElement.Comment = dataReader[0].ToString();
                            dBElement.Description = dataReader[1].ToString();
                            dBElement.PartNumber = dataReader[2].ToString();
                            dBElement.TU = dataReader[3].ToString();
                            if (dataReader[4] != null && dataReader[4].ToString() != string.Empty)
                            {
                                var range = dataReader[4].ToString();
                                Regex regexRange = new Regex(@"пФ|мкФ|Ом|кОм|МОм");
                                if (regexRange.Matches(range).Count > 1  || !range.Contains(";"))
                                {
                                    dBElement.RangeValue = range.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                }
                                else
                                {
                                    //Если размерность для диапазона указана в конце списка, добавлем ее каждому номиналу
                                    dBElement.RangeValue = range.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                    for (int i = 0; i < dBElement.RangeValue.Count; i++)
                                    {
                                        if (!dBElement.RangeValue[i].Contains(regexRange.Match(range).Value.ToString()))
                                        dBElement.RangeValue[i] = dBElement.RangeValue[i] + " " + regexRange.Match(range).Value.ToString();
                                    }
                                }
                            }
                            dBElement.OptionalParameter = dataReader[5].ToString();
                            dBElement.Admission = dataReader[6].ToString();
                            dBElement.Value = dataReader[7].ToString();
                        }

                        if (dataReader[8] != null && dataReader[8].ToString() != string.Empty)
                        {
                            linkedObject.CodePKI = dataReader[8].ToString();

                            linkedObject.Count = Convert.ToInt32(dataReader[9]);
                            dBElement.LinkedObject.Add(linkedObject);
                            countLinkedObj++;
                        }

                        //    if (dataReader[3] != null && dataReader[3].ToString() != null)
                        //   dBElement.RangeValue = 
                        prewElement = dataReader[2].ToString();
                        dBElements.Add(dBElement);
                    }

                }

                //    MessageBox.Show(dBElements.Count + " " + countLinkedObj);

                //var linkedCount = 0;
                //var dBElementsDistinct = new List<DataBaseElements>();

                //foreach (var obj in dBElements)
                //{
                //    if (dBElementsDistinct.Where(t => t.PartNumber == obj.PartNumber).ToList().Count == 0)
                //    {
                //        dBElementsDistinct.Add(obj);
                //    }
                //}

                //   StringBuilder strb = new StringBuilder();

                int linkedCount = 0;
                List<LinkedObjectDB> linkedObjects = new List<LinkedObjectDB>();



                OleDbCommand cmdLinkedObjects = con.CreateCommand();

                var queryLinkObjects = string.Format(@"SELECT ID,
                                                          [Обозначение], 
                                                          [Наименование],
                                                          [Заготовка],
                                                          [Примечание],
                                                          [Тип]     
                                                   FROM [Связанные объекты]");

                cmdLinkedObjects.CommandText = queryLinkObjects;
                cmdLinkedObjects.Connection = con;
                OleDbDataReader dataReaderLinkedObjects = cmdLinkedObjects.ExecuteReader();

                while (dataReaderLinkedObjects.Read())
                {
                    LinkedObjectDB linkedObject = new LinkedObjectDB();
                    linkedCount++;
                    linkedObject.Id = Convert.ToInt32(dataReaderLinkedObjects[0].ToString().Trim());
                    linkedObject.Denotation = dataReaderLinkedObjects[1].ToString();
                    linkedObject.Name = dataReaderLinkedObjects[2].ToString();

                    linkedObject.Preparation = dataReaderLinkedObjects[3].ToString();
                    linkedObject.Comment = dataReaderLinkedObjects[4].ToString();
                    linkedObject.Type = dataReaderLinkedObjects[5].ToString();
                    linkedObjects.Add(linkedObject);
                }

                linkedObjects = linkedObjects.Where(t => t.Name != string.Empty).ToList();
                //    reportParameters.LinkedObjects.AddRange(linkedObjects);

                //  MessageBox.Show(linkedObjects.Count + "");


                foreach (var element in dBElements.Distinct())
                {
                    foreach (var linkedObject in element.LinkedObject)
                    {
                        foreach (var obj in linkedObjects)
                        {

                            if (Convert.ToInt32(linkedObject.CodePKI.Trim()) == obj.Id)
                            {
                                linkedObject.Denotation = obj.Denotation;
                                linkedObject.Name = obj.Name;
                                linkedObject.Preparation = obj.Preparation;
                                linkedObject.Comment = obj.Comment;
                                linkedObject.Type = obj.Type;
                            }
                        }
                    }
                }

                //if (linkedObjects.Where(t=>t.Denotation == "ГН7.840.835").ToList().Count > 0)
                //{
                //    MessageBox.Show("Нашел Count - " + linkedObjects.Where(t => t.Denotation == "ГН7.840.835").ToList().Count +" CodePKI" + linkedObjects.Where(t => t.Denotation == "ГН7.840.835").ToList()[0].Id);
                //}

                //StringBuilder strBuilding = new StringBuilder();

                //foreach (var element in dBElements.Distinct())
                //{
                //    foreach (var obj in element.LinkedObject.OrderBy(t => t.CodePKI).ToList())
                //    {
                //        strBuilding.AppendLine(obj.CodePKI + " " + obj.Denotation + " " + obj.Name + " " + obj.Type);
                //    }
                //}
                //MessageBox.Show(strBuilding.ToString());
                //Clipboard.SetText(strBuilding.ToString());



                //try
                //{
                //    foreach (var element in dBElements.Distinct())
                //    {
                //        foreach (var linkedObject in element.LinkedObject)
                //        {
                //            k = 0;
                //            OleDbCommand cmdLinkedObjects = con.CreateCommand();
                //            k = 1;
                //            var query = string.Format(@"SELECT [Обозначение], 
                //                                           [Наименование],
                //                                           [Заготовка],
                //                                           [Примечание],
                //                                           [Тип]     
                //                                   FROM [Связанные объекты]
                //                                   WHERE ID = '{0}'", linkedObject.CodePKI);
                //            k = 2;
                //            cmdLinkedObjects.CommandText = query;
                //            k = 3;
                //          //  cmd.Connection = con;
                //            k = 4;
                //            OleDbDataReader dataReaderLinkedObjects = cmd.ExecuteReader();
                //            k = 5;
                //            while (dataReaderLinkedObjects.Read())
                //            {
                //                k = 6;
                //                linkedCount++;
                //                linkedObject.Denotation = dataReader[0].ToString();
                //                linkedObject.Name = dataReader[1].ToString();
                //                linkedObject.Preparation = dataReader[2].ToString();
                //                linkedObject.Comment = dataReader[3].ToString();
                //                linkedObject.Type = dataReader[4].ToString();
                //              //  strb.AppendLine(element.PartNumber + " -   " + linkedObject.Denotation + " " + linkedObject.Name + " " + linkedObject.Comment + " " + linkedObject.Type);
                //            }
                //        }
                //    }
                //}
                //catch
                //{
                //    MessageBox.Show("пиздец "+ linkedCount + " " + k);
                //}

                ///  MessageBox.Show("успех " + linkedCount + "");

                //foreach (var element in dBElements)
                //{
                //    foreach (var linkedObject in element.LinkedObject)
                //    {
                //    }
                //}



                //        OleDbCommand cmd = con.CreateCommand();
                //        var query = string.Format(@"SELECT [Обозначение], 
                //                                           [Наименование],
                //                                           [Заготовка],
                //                                           [Примечание],
                //                                           [Тип]     
                //                                   FROM [Связанные объекты]
                //                                   WHERE ID IN '{0}'", linkedObject.CodePKI);
                //        cmd.CommandText = query;
                //        cmd.Connection = con;
                //        OleDbDataReader dataReader = cmd.ExecuteReader();

                //while (dataReader.Read())
                //{
                //    // linkedCount++;
                //    linkedObject.Denotation = dataReader[0].ToString();
                //    linkedObject.Name = dataReader[1].ToString();
                //    linkedObject.Preparation = dataReader[2].ToString();
                //    linkedObject.Comment = dataReader[3].ToString();
                //    linkedObject.Type = dataReader[4].ToString();
                //    //       strb.AppendLine(element.PartNumber + " -   " + linkedObject.Denotation + " " + linkedObject.Name + " " + linkedObject.Comment + " " + linkedObject.Type);
                //}



                // MessageBox.Show("Количество присоединенных объектов " + linkedCount + " " + dBElementsDistinct.Count + " " + dBElements.Count);
              //  con.Close();
                // MessageBox.Show(strb.ToString());
            }
        }
        private void checkBoxApprovedBy_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxApprovedBy.CheckState == CheckState.Checked)
                dTPApprovedBy.Visible = true;
            else
                dTPApprovedBy.Visible = false;
        }

        private void checkBoxNControllBy_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxNControllBy.CheckState == CheckState.Checked)
                dTPNControlBy.Visible = true;
            else
                dTPNControlBy.Visible = false;
        }

        private void checkBoxMasteredBy_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMasteredBy.CheckState == CheckState.Checked)
                dTPMasteredBy.Visible = true;
            else
                dTPMasteredBy.Visible = false;
        }

        private void checkBoxAuthor_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAuthor.CheckState == CheckState.Checked)
                dTPAuthor.Visible = true;
            else
                dTPAuthor.Visible = false;
        }

        private void radioButtonPCAD_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonPCAD.Checked)
            {
                openFileDialog.Filter = "Файлы ATR|*.atr";
            }
            else
            {
                openFileDialog.Filter = "Файлы XLSX|*.xlsx; *.xls";
            }
        }

        private void radioButtonAltiumDesigner_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonAltiumDesigner.Checked)
            {
                openFileDialog.Filter = "Файлы XLS|*.xlsx; *.xls";
            }
            else
            {
                openFileDialog.Filter = "Файлы ATR|*.atr";
            }
        }

        private void checkBoxLoadKre_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBoxFileName_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
