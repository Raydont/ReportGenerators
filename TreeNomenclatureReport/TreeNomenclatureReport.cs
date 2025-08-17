using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using TFlex.Reporting;
using ReportHelpers;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.UI.Objects.Managers;
using TFlex.DOCs.Model.Plugins;
using TFlex.DOCs.Model.Macros;

namespace TreeNomenclatureReport
{
    public class ExcelGenerator : IReportGenerator
    {
        /// <summary>
        /// Расширение файла отчета
        /// </summary>
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xls";
            }
        }

        /// <summary>
        /// Определяет редактор параметров отчета ("Параметры шаблона" в контекстном меню Отчета)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="context"></param>
        /// <param name="owner"></param>
        /// <returns></returns>
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            context.CopyTemplateFile();    // Создаем копию шаблона
            bool isCatch = false;

            Xls xls = new Xls();

            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);
                var report = new Report();
                report.MakeNomTreeReport(context, xls);
                xls[1, 1].Select();
                xls.AutoWidth();
                xls.Save();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString(), "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (isCatch)
                    xls.Quit(false);
                else
                {
                    xls.Quit(true);
                    // Открытие файла
                    System.Diagnostics.Process.Start(context.ReportFilePath);
                }
            }
        }
    }

    class NomParams
    {
        public int id;
        public int parentId;
        public int level;
        public string name;
        public string denotation;
        public string unitCode;
        public double amount;
        public string bomSection;
    }

    public class Report
    {
        List<NomParams> tree = new List<NomParams>();
        /*
                private List<NomParams> ReadData(int objectId)
                {
                    // соединяемся с БД T-FLEX DOCs
                    SqlConnection conn;
                    SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
                    sqlConStringBuilder.DataSource = "S2";
                    sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
                    sqlConStringBuilder.Password = "reportUser";
                    sqlConStringBuilder.UserID = "reportUser";
                    conn = new SqlConnection(sqlConStringBuilder.ToString());
                    conn.Open();

                    #region Запрос считывания данных заказа
                    string getNomTreeCommandQuery = String.Format(@"
                                    DECLARE @docid INT
                                    DECLARE @productId INT
                                    DECLARE @level INT
                                    DECLARE @insertcount INT
                                    DECLARE @CollectionID INT
                                    SET @docid = {0}
                                    SET @insertcount = 0
                                    SET @level=0

                                    IF OBJECT_ID('tempdb..#TmpR') IS NULL
                                    CREATE TABLE #TmpR (s_ObjectID INT,             
                                                          s_ParentID INT, 
                                                          [level] INT,
                                                          s_ClassID INT,                      
                                                          Name NVARCHAR(255),
                                                          Denotation NVARCHAR(150),
                                                          Amount FLOAT,
                                                          BomSection NVARCHAR(255),
                                                          UnitCode NVARCHAR(255) NULL
                                                          --,TotalCount FLOAT
                                                          ) 
                                    ELSE TRUNCATE TABLE #TmpR

                                    INSERT INTO #TmpR
                                    SELECT DISTINCT n.s_ObjectID, -1, 0, n.s_ClassID,
                                            n.Name, n.Denotation, 1, 
                                            N'Комплексы', duc.Code
                                            --,1
                                    FROM Nomenclature n INNER JOIN Documents d ON (d.s_NomenclatureObjectID = n.s_ObjectID)
                                                        LEFT JOIN UnitCodes duc ON (n.s_ObjectID = duc.s_ObjectID)                                       
                                    WHERE n.s_ObjectID=@docid AND
                                          n.s_Deleted=0 AND n.s_ActualVersion=1 AND
                                          d.s_Deleted = 0 AND d.s_ActualVersion = 1 AND    
                                          duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 
    
                                    WHILE 1=1
                                    BEGIN

                                       INSERT INTO #TmpR
                                       SELECT DISTINCT nh.s_ObjectID, nh.s_ParentID, @level+1,
                                              n.s_ClassID, n.Name, n.Denotation,
                                              nh.Amount, nh.BOMSection, ISNULL(duc.Code,'') as UnitCode
                                              -- ,#TmpR.TotalCount*nh.Amount
                                       FROM NomenclatureHierarchy nh INNER JOIN #TmpR ON nh.s_ParentID = #TmpR.s_ObjectID AND
                                                                                           #TmpR.[level] = @level AND
                                                                                           nh.s_ActualVersion = 1 AND nh.s_Deleted = 0
                                                                     INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID AND
                                                                                                  n.s_ActualVersion = 1 AND n.s_Deleted = 0 AND
                                                                                                  n.s_ClassID IN (500, 501, 566)--, 510, 839,499)
                                                                                                  -- Сборочная единица, Комплект, 
                                                                                                  -- Прочее изделие, Стандартное изделие, Материалы
                                                                     LEFT JOIN Documents d ON nh.s_ObjectID = d.s_NomenclatureObjectID AND
                                                                                              d.s_Deleted=0 AND d.s_ActualVersion=1
                                                                     LEFT JOIN UnitCodes duc ON n.s_ObjectID = duc.s_ObjectID AND
                                                                                                duc.s_Deleted = 0 AND duc.s_ActualVersion = 1
                                
                                        SET @insertcount = @@ROWCOUNT 
                                        SET @level = @level + 1

                                        IF  @insertcount = 0
                                        GOTO END1
                                    END             	
                                    END1:

                                    SELECT * FROM #TmpR", objectId);
                    #endregion

                    SqlCommand docTreeCommand = new SqlCommand(getNomTreeCommandQuery, conn);
                    docTreeCommand.CommandTimeout = 0;
                    SqlDataReader reader = docTreeCommand.ExecuteReader();

                    while (reader.Read())
                    {
                        var treeItem = new NomParams();
                        treeItem.id = reader.GetInt32(0);
                        treeItem.parentId = reader.GetInt32(1);
                        treeItem.level = reader.GetInt32(2);
                        treeItem.name = reader.GetString(4);
                        treeItem.denotation = reader.GetString(5);
                        treeItem.amount = reader.GetDouble(6);
                        treeItem.bomSection = reader.GetString(7);
                        treeItem.unitCode = reader.IsDBNull(8) ? "" : reader.GetString(8);

                        tree.Add(treeItem);
                    }

                    return tree;
                }*/

        public static readonly int AssemblyClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Сборочная единица").Id;
        public static readonly int DetalClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Деталь").Id;
        public static readonly int DocumentClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find(new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114")).Id;
        public static readonly int KomplektClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Комплект").Id;
        public static readonly int ComplexClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Комплекс").Id;
        public static readonly int OtherItemClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Прочее изделие").Id;
        public static readonly int StandartItemClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Стандартное изделие").Id;
        public static readonly int MaterialClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find(new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987")).Id;
        public static readonly int ComponentProgramClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Компонент (программы)").Id;
        public static readonly int ComplexProgramClassId = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Комплекс (программы)").Id;

        public void MakeNomTreeReport(IReportGenerationContext context, Xls xls)
        {
            TFlex.DOCs.UI.Objects.TFlexDOCsClientUI.Initialize();
            var dialog = ObjectCreator.CreateObject<IInputDialog>();
            dialog.Caption = "Дерево состава";

            dialog.AddFlagField("Документация", false, false, false);
            dialog.AddFlagField("Комплексы", false, false, false);
            dialog.AddFlagField("Детали", false, false, false);
            dialog.AddFlagField("Сборочные единицы", true, false, false);
            dialog.AddFlagField("Стандартные изделия", false, false, false);
            dialog.AddFlagField("Прочие изделия", true, false, false);
            dialog.AddFlagField("Материалы", false, false, false);
            dialog.AddFlagField("Комплекты", true, false, false);
            dialog.AddFlagField("Комплексы (программы)", false, false, false);
            dialog.AddFlagField("Компоненты (программы)", false, false, false);

            dialog.FieldValueChanged += Dialog_FieldValueChanged; ;

            dialog.Show(new MacroContext(ServerGateway.Connection));

            int objectId;
            var originalTree = new List<NomParams>();

            // получаем ID выделенного в интерфейсе T-FLEX DOCs объекта
            if (context.ObjectsInfo.Count == 1) objectId = context.RootObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return;
            }

            var reference = ServerGateway.Connection.ReferenceCatalog.Find(context.ReferenceId).CreateReference();
            var rootObject = reference.Find(objectId);
            var hierarchy = rootObject.Children.RecursiveLoadHierarchyLinks();
            var fullHierarchy = new Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>>();
            var allObjects = new Dictionary<int, ReferenceObject>();
            allObjects[objectId] = rootObject;
            foreach (var link in hierarchy)
            {
                allObjects[link.ChildObjectId] = link.ChildObject;
                Dictionary<Guid, ComplexHierarchyLink> objectHierarchy;
                if (!fullHierarchy.TryGetValue(link.ParentObjectId, out objectHierarchy))
                {
                    objectHierarchy = new Dictionary<Guid, ComplexHierarchyLink>();
                    fullHierarchy[link.ParentObjectId] = objectHierarchy;
                }
                objectHierarchy[link.Guid] = link;
            }

            FillChildrenNew(objectId, 1, 1, fullHierarchy, allObjects, originalTree);

            /*
            var treeData = ReadData(objectId);

            if (treeData.Count == 0)
            {
                MessageBox.Show("0");
                return;
            }

            var root = treeData[0];

            FillChildren(root, originalTree, treeData);*/

            Stack<int> startLines = new Stack<int>();
            int lastLevel = -1;
            var elements = originalTree.Select(t => t.level).ToList();
            var width = (elements.Count>0? elements.Max():0) + 4;

            // Группировка сверху-вниз
            xls.Worksheet.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;

            for (int i = 0; i < originalTree.Count; i++)
            {
                var item = originalTree[i];
                var col = 2;
                var row = i + 1;

                xls[col++, row].SetValue(new string(' ', item.level * 4) + item.name);
                xls[col++, row].SetValue(item.denotation);
                xls[col++, row].SetValue(item.amount);
                xls[col + 3, row].SetValue(item.id);

                if (lastLevel != item.level)
                {
                    if (item.level > lastLevel)
                    {
                        startLines.Push(row);
                    }
                    else
                    {
                        var countLevels = lastLevel - item.level;
                        for (int j = 0; j < countLevels; j++)
                        {
                            var startGroupLine = startLines.Pop();
                            var endGroupLine = row - 1;

                            // Группировка данных
                            xls[1, startGroupLine, width, endGroupLine - startGroupLine + 1].Rows.Group();
                            //xls[1, startGroupLine, width, endGroupLine - startGroupLine + 1].EntireRow.Hidden = false;
                        }
                    }
                    lastLevel = item.level;
                }
            }

            while (startLines.Count > 2)
            {
                var startGroupLine = startLines.Pop();
                var endGroupLine = originalTree.Count;

                // Группировка данных
                xls[1, startGroupLine, width, endGroupLine - startGroupLine + 1].Rows.Group();
                //xls[1, startGroupLine, width, endGroupLine - startGroupLine + 1].EntireRow.Hidden = false;
            }
            xls[1, 1, 5, originalTree.Count + 10].Columns.AutoFit();

            xls.Worksheet.Outline.ShowLevels(1);

            //xls.Visible = true;
        }

        private void Dialog_FieldValueChanged(object sender, FieldValueChangedEventArgs e)
        {
            var flag = (bool)e.NewValue;
            switch (e.Name)
            {
                case "Документация":
                    supportedClasses.Remove(DocumentClassId);
                    if (flag) supportedClasses.Add(DocumentClassId);
                    break;
                case "Комплексы":
                    supportedClasses.Remove(ComplexClassId);
                    if (flag) supportedClasses.Add(ComplexClassId);
                    break;
                case "Детали":
                    supportedClasses.Remove(DetalClassId);
                    if (flag) supportedClasses.Add(DetalClassId);
                    break;
                case "Сборочные единицы":
                    supportedClasses.Remove(AssemblyClassId);
                    if (flag) supportedClasses.Add(AssemblyClassId);
                    break;
                case "Стандартные изделия":
                    supportedClasses.Remove(StandartItemClassId);
                    if (flag) supportedClasses.Add(StandartItemClassId);
                    break;
                case "Прочие изделия":
                    supportedClasses.Remove(OtherItemClassId);
                    if (flag) supportedClasses.Add(OtherItemClassId);
                    break;
                case "Материалы":
                    supportedClasses.Remove(MaterialClassId);
                    if (flag) supportedClasses.Add(MaterialClassId);
                    break;
                case "Комплекты":
                    supportedClasses.Remove(KomplektClassId);
                    if (flag) supportedClasses.Add(KomplektClassId);
                    break;
                case "Комплексы (программы)":
                    supportedClasses.Remove(ComplexProgramClassId);
                    if (flag) supportedClasses.Add(ComplexProgramClassId);
                    break;
                case "Компоненты (программы)":
                    supportedClasses.Remove(ComponentProgramClassId);
                    if (flag) supportedClasses.Add(ComponentProgramClassId);
                    break;
            }
        }

        List<int> supportedClasses = new List<int> { AssemblyClassId, OtherItemClassId, KomplektClassId };

        private void FillChildrenNew(int objectId,
            int level,
            double count,
            Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>> fullHierarchy,
            Dictionary<int, ReferenceObject> allObjects,
            List<NomParams> originalTree)
        {

            var refObj = (NomenclatureObject)allObjects[objectId];
            if (!supportedClasses.Contains(refObj.Class.Id)) return;

            var root = new NomParams();
            originalTree.Add(root);
            root.level = level;
            root.id = objectId;
            root.denotation = refObj.Denotation;
            root.name = refObj.Name;
            root.amount = count;

            if (fullHierarchy.ContainsKey(objectId))
            {
                var children = fullHierarchy[objectId].Values
                    .OrderBy(t => ((NomenclatureObject)t.ChildObject).Denotation.GetString())
                    .ThenBy(t => ((NomenclatureObject)t.ChildObject).Name.GetString());
                foreach (var child in children)
                {
                    var nomLink = (NomenclatureHierarchyLink)child;
                    FillChildrenNew(child.ChildObjectId, level + 1, nomLink.Amount.GetDouble(), fullHierarchy, allObjects, originalTree);
                }
            }
        }
        /*
        private void FillChildren(NomParams root, List<NomParams> originalTree, List<NomParams> treeData)
        {
            originalTree.Add(root);
            var children = treeData.Where(t => t.level == root.level + 1 && t.parentId == root.id).ToList();
            foreach (var child in children)
            {
                FillChildren(child, originalTree, treeData);
            }
        }*/
    }
}
