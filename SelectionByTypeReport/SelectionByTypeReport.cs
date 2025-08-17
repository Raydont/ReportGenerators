using Microsoft.Office.Interop.Excel;
using ReportHelpers;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using Font = System.Drawing.Font;

namespace SelectionByTypeReport
{
    public class ExcelGenerator : IReportGenerator
    {
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

        //Создание экземпляра формы выбора параметров формирования отчета
        public SelectionForm selectionForm = new SelectionForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            using (new WaitCursorHelper(false))
            {
                var vpmoDepartmentGuid = new Guid("28c1d4b9-8c74-4ff1-a336-e1756ae04663");
                var chiefVpmo = new Guid("57672647-f64d-4da8-aeed-cf5b1729aad2");
                var employesVpmo = new Guid("299f0f75-e163-48a8-8efd-4d79fa3b2e85");
                var vpmoGroup = new List<Guid>()
                {
                    vpmoDepartmentGuid,
                    chiefVpmo,
                    employesVpmo
                };
                
                var currentUser = ServerGateway.Connection.ClientView.GetUser().Parents.GetObjects();

                if (vpmoGroup.Any(t => currentUser.Select(p => p.SystemFields.Guid).Contains(t)))
                {
                    MessageBox.Show("В связи с окончанием поддержки версии T-FLEX DOCs 15 и переходом на новую версию T-FLEX DOCs 17 с новым API отчет Выборка по типам работает некорректно. Воспользуйтесь другим отчетом.",
                        "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    selectionForm.Init(context);
                    if (selectionForm.ShowDialog() == DialogResult.OK)
                    {
                        selectionForm.MakeReport();
                    }
                }
            }
        }
    }

    public class WaitCursorHelper : IDisposable
    {
        private bool _waitCursor;

        public WaitCursorHelper(bool useWaitCursor)
        {
            _waitCursor = System.Windows.Forms.Application.UseWaitCursor;
            System.Windows.Forms.Application.UseWaitCursor = useWaitCursor;
        }

        public void Dispose()
        {
            System.Windows.Forms.Application.UseWaitCursor = _waitCursor;
        }
    }

    // ТЕКСТ МАКРОСА ===================================

    public delegate void LogDelegate(string line);

    public class ReportParameters
    {
        //поля отчета
        bool position; //Позиция
        bool zone; //Зона   
        bool format; //Формат   
        bool denotation; //Обозначение
        bool name; //Наименование
        bool count; //количество
        bool remarks; //примечание
        bool id; // id элемента
        bool letter; // литера 
        bool codeRCP; // Код ОКП
        bool firstUse; // Первичная применяемость
        bool createDate; // Первичная применяемость
        bool nomenclatureNumber; //Номенклатурный номер
        bool measureUnit; //единица измерения



        // типы объектов
        bool documentation; //Документация
        bool complex; //Комплексы
        bool detail; //Детали
        bool assembly; //Сборочные единицы
        bool standart; //Стандартные изделия
        bool other; //Прочие изделия
        bool materials; //Материалы
        bool complement; //Комплекты
        bool complexProgram; //Комплексы (программы)
        bool componentProgram; //Компоненты (программы) 

        public string LoadFolderName;
        public bool WithOutKooperationDSE;

        //типы объектов для Вывода отчета о наличии техпроцессов
        public bool DetailTP = false;
        public bool ComplementTP = false;
        public bool AssemblyTP = false;
        public bool ComplexTP = false;
        public bool StandartItemsTP = false;

        public bool GroupByDevice = false;
        public bool GroupByTypes = false;
        public bool GroupByDeviceAndTypes = false;


        public bool OutputTechProcesses = false;

        // тип отчета
        bool expansionNodes;

        // отчет по 5-ти основным полям 
        public bool Main5Col;

        // список удаляемых объектов из состава объекта отчета
        public List<int> listDeleteObjectId;

        public Dictionary<int, int> listObjectCountId;

        // параметр добавление условия содержит наименование
        public List<string> AddLikeName = new List<string>();

        // параметр добавление условия содержит обозначение
        public List<string> AddLikeDenotation = new List<string>();

        public bool Position
        {
            get { return position; }
            set { position = value; }
        }

        public bool Zone
        {
            get { return zone; }
            set { zone = value; }
        }

        public bool Format
        {
            get { return format; }
            set { format = value; }
        }

        public bool Denotation
        {
            get { return denotation; }
            set { denotation = value; }
        }

        public bool Name
        {
            get { return name; }
            set { name = value; }
        }

        public bool Count
        {
            get { return count; }
            set { count = value; }
        }

        public bool Remarks
        {
            get { return remarks; }
            set { remarks = value; }
        }

        public bool Letter
        {
            get { return letter; }
            set { letter = value; }
        }

        public bool Id
        {
            get { return id; }
            set { id = value; }
        }

        public bool AuthorName;
        public bool LastEditorName;

        public bool Documentation
        {
            get { return documentation; }
            set { documentation = value; }
        }

        public bool Complex
        {
            get { return complex; }
            set { complex = value; }
        }

        public bool Detail
        {
            get { return detail; }
            set { detail = value; }
        }

        public bool Assembly
        {
            get { return assembly; }
            set { assembly = value; }
        }

        public bool Standart
        {
            get { return standart; }
            set { standart = value; }
        }

        public bool Other
        {
            get { return other; }
            set { other = value; }
        }

        public bool Materials
        {
            get { return materials; }
            set { materials = value; }
        }

        public bool Complement
        {
            get { return complement; }
            set { complement = value; }
        }

        public bool ComplexProgram
        {
            get { return complexProgram; }
            set { complexProgram = value; }
        }

        public bool ComponentProgram
        {
            get { return componentProgram; }
            set { componentProgram = value; }
        }

        public bool ExpansionNodes
        {
            get { return expansionNodes; }
            set { expansionNodes = value; }
        }

        public bool CodeRcp
        {
            get { return codeRCP; }
            set { codeRCP = value; }
        }

        public bool FirstUse
        {
            get { return firstUse; }
            set { firstUse = value; }
        }

        public bool CreateDate
        {
            get { return createDate; }
            set { createDate = value; }
        }

        public bool NomenclatureNumber
        {
            get { return nomenclatureNumber; }
            set { nomenclatureNumber = value; }
        }

        public bool MeasureUnit { get => measureUnit; set => measureUnit = value; }
    }

    public class TFDDocument
    {
        // Параметры объекта

        public int ObjectID;
        public int ParentID;
        public int Class;
        public double Amount;
        public List<string> Zone = new List<string>();
        public List<string> Position = new List<string>();
        public string Format;
        public string Naimenovanie;
        public string Denotation;
        public List<string> Remarks = new List<string>();
        public int Level;
        public string Letter;
        public string CodeRCP;
        public string FirstUse;
        public string CreateDate;
        public string NomenclatureNumber;
        public string AuthorName;
        public string LastEditorName;
        public List<string> MeasureUnit = new List<string>();
        public ReferenceObject ParentObject = null;

        public int parenDevicetId;

        public List<TechnologicalProcess> TechProcesses = new List<TechnologicalProcess>();


        public NomenclatureObject NomObj;

        public TFDDocument()
        {
        }

        public TFDDocument(NomenclatureObject nomObject, ComplexHierarchyLink hierarchy, ReportParameters reportParameters)
        {
            NomObj = nomObject;
            if (reportParameters.Id)
                ObjectID = nomObject.SystemFields.Id;

            if (reportParameters.Name)
                Naimenovanie = nomObject[Guids.Nomenclature.Fields.Name].Value.ToString();

            if (reportParameters.Denotation)
                Denotation = nomObject[Guids.Nomenclature.Fields.Denotation].Value.ToString();

            if (reportParameters.Format)
                Format = nomObject[Guids.Nomenclature.Fields.Format].Value.ToString();

            if (reportParameters.CreateDate)
                CreateDate = nomObject.SystemFields.CreationDate.ToShortDateString();

            if (reportParameters.AuthorName)
                AuthorName = nomObject.SystemFields.AuthorName;

            if (reportParameters.LastEditorName)
                AuthorName = nomObject.SystemFields.EditorName;

            Class = nomObject.Class.Id;

            if (hierarchy != null && hierarchy.ParentObject != null)
                ParentObject = hierarchy.ParentObject;
            try
            {
                if (reportParameters.FirstUse)
                    FirstUse = nomObject[Guids.Nomenclature.Fields.FirstUse].Value.ToString();
            }
            catch
            {

            }
            try
            {
                if (reportParameters.CodeRcp)
                    CodeRCP = nomObject[Guids.Nomenclature.Fields.CodeRCP].Value.ToString();
            }
            catch
            {

            }



            if (reportParameters.Remarks)
                Remarks.Add(SetStringHierarhyParameters(hierarchy, Guids.NomHierarchyLink.Fields.Comment));

            if (reportParameters.MeasureUnit)
                MeasureUnit.Add(SetStringHierarhyParameters(hierarchy, Guids.NomHierarchyLink.Fields.MeasureUnit));

            if (reportParameters.Zone)
                Zone.Add(SetStringHierarhyParameters(hierarchy, Guids.NomHierarchyLink.Fields.Zone));

            if (reportParameters.Position)
                Position.Add(SetStringHierarhyParameters(hierarchy, Guids.NomHierarchyLink.Fields.Pos, true));

        }

        public TFDDocument(NomenclatureObject nomObject, ReportParameters reportParameters)
        {
            NomObj = nomObject;
            if (reportParameters.Id)
                ObjectID = nomObject.SystemFields.Id;

            if (reportParameters.Name)
                Naimenovanie = nomObject[Guids.Nomenclature.Fields.Name].Value.ToString();

            if (reportParameters.Denotation)
                Denotation = nomObject[Guids.Nomenclature.Fields.Denotation].Value.ToString();

            if (reportParameters.Format)
                Format = nomObject[Guids.Nomenclature.Fields.Format].Value.ToString();

            if (reportParameters.CreateDate)
                CreateDate = nomObject.SystemFields.CreationDate.ToShortDateString();

            if (reportParameters.AuthorName)
                AuthorName = nomObject.SystemFields.AuthorName;

            if (reportParameters.LastEditorName)
                AuthorName = nomObject.SystemFields.EditorName;

            Class = nomObject.Class.Id;

        }

        public TFDDocument(NomenclatureObject nomObject)
        {
            NomObj = nomObject;
            ObjectID = nomObject.SystemFields.Id;
            Naimenovanie = nomObject[Guids.Nomenclature.Fields.Name].Value.ToString();
            Denotation = nomObject[Guids.Nomenclature.Fields.Denotation].Value.ToString();
        }


        public bool Equals(TFDDocument doc)
        {
            if (doc == null)
                return false;
            else
            {
                var equal = NomObj == doc.NomObj &&
                //Параметры иерархии
                (Remarks.SequenceEqual(doc.Remarks) || !doc.NomObj.Class.IsInherit(Guids.Nomenclature.Types.Material)) &&       //Проверяем примечание, только у материалов
                (MeasureUnit.SequenceEqual(doc.MeasureUnit) || !doc.NomObj.Class.IsInherit(Guids.Nomenclature.Types.Material)); //Проверяем единицу измерения только у материалов

                //Зону и позицию не проверяем
                //Zone == doc.Zone &&
                //Position == doc.Position;

                //Если объекты идентичны, а параметры иерархии различны, то добавляем различные параметры к sourse объекту
                if (equal)
                {
                    if (!doc.Remarks.SequenceEqual(Remarks))
                    {
                        Remarks.AddRange(doc.Remarks.Where(t => !Remarks.Contains(t)));
                    }

                    if (!doc.MeasureUnit.SequenceEqual(MeasureUnit))
                    {
                        MeasureUnit.AddRange(doc.MeasureUnit.Where(t => !MeasureUnit.Contains(t)));
                    }

                    if (!doc.Zone.SequenceEqual(Zone))
                    {
                        Zone.AddRange(doc.Zone.Where(t => !Zone.Contains(t)));
                    }

                    if (!doc.Position.SequenceEqual(Position))
                    {
                        Position.AddRange(doc.Position.Where(t => !Position.Contains(t)));
                    }
                }

                return equal;
            }
        }

        public string ToString()
        {
            return "[" + ObjectID + "] " + Naimenovanie + " " + Denotation + " " + NomObj + " " + string.Join(" ", Remarks) + " " + string.Join(" ", MeasureUnit) + " " + string.Join(" ", Zone) + " " + string.Join(" ", Position);
        }



        public string SetStringHierarhyParameters(ComplexHierarchyLink hierarchy, Guid parametrGuid, bool intValue = false)
        {
            var parHierarchy = string.Empty;

            if (hierarchy != null)
            {
                if (intValue)
                    parHierarchy = hierarchy[parametrGuid].GetInt32().ToString().Trim();
                else
                    parHierarchy = hierarchy[parametrGuid].GetString().Trim();
            }
            return parHierarchy;
        }
    }

    public class TechnologicalProcess
    {
        //дополнительные поля техпроцессов
        public string Name;               //Наименование техпроцесса
        public double? SumPieceTime;       //Сумма Tшт
        public double? SumPrepTim;         //Сумма Тпз
        public double? PieceTime;          //Штучное время 
        public string OperName;           //Наименование Операции
        public string Remarks;            //Примечание
        public int? ObjectID;
    }

    public class SelectionByTypeReport : IDisposable
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters, Action<List<TFDDocument>, IndicationForm, LogDelegate> getDocuments)
        {
            SelectionByTypeReport report = new SelectionByTypeReport();
            report.Make(context, reportParameters, getDocuments);
        }

        public void Make(IReportGenerationContext context, ReportParameters reportParameters, Action<List<TFDDocument>, IndicationForm, LogDelegate> getDocuments = null)
        {
            if (!reportParameters.LoadFolderName.Contains("Выгрузка"))
                context.CopyTemplateFile();    // Создаем копию шаблона

            m_form.TopMost = true;
            m_form.Visible = true;
            m_form.Activate();
            LogDelegate m_writeToLog;
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            List<TFDDocument> reportData = new List<TFDDocument>();
            List<TFDDocument> sortedReportData;

            Xls xls = new Xls();
            bool isCatch = false;

            SelectionByTypeReport report = new SelectionByTypeReport();

            if (reportParameters.listObjectCountId.Count == 0)
            {
                MessageBox.Show("Выберите объект для формирования отчета");
                return;
            }

            m_form.setStage(IndicationForm.Stages.DataAcquisition);
            m_writeToLog("=== Получение данных ===");

            //Формируем список из справочника ДСЕ по кооперации

            if (reportParameters.WithOutKooperationDSE)
            {
                List<ReferenceObject> dseKooperation = new List<ReferenceObject>();
                dseKooperation.AddRange(GetDSEKooperation());

                foreach (var id in reportParameters.listObjectCountId)
                {
                    //Проверяем встречается ли сборка ДСЕ по кооперации в основном списке, если да и флаг с учетом кооперации выставлен пропускаем эту сборку
                    if (dseKooperation.Count == 0 || dseKooperation.Count(t => t.SystemFields.Id == id.Key) == 0)
                    {
                        reportData.AddRange(ReadData(id, m_writeToLog, reportParameters, dseKooperation));
                    }
                }

                // Clipboard.SetText(string.Join("\r\n", reportData.Select(t => t.NomObj + " " + t.Amount)));
                //  MessageBox.Show("1 " + string.Join("\r\n", reportData.Select(t => t.NomObj + " " + t.Amount)));

                //Повторная чистка от ДСЕ внутри других сборок
                foreach (var dse in dseKooperation)
                {
                    if (dse.Class.IsInherit(Guids.Nomenclature.Types.Detail))
                    {
                        reportData = reportData.Where(t => t.NomObj.SystemFields.Id != dse.SystemFields.Id).ToList();
                    }

                    /* if (dse.Class.IsInherit(Guids.Nomenclature.Types.Assembly))
                     {
                         var dseAssembly = reportData.Where(t => t.NomObj.SystemFields.Id == dse.SystemFields.Id).FirstOrDefault();
                         if (dseAssembly != null)
                         {
                             var dseId = new KeyValuePair<int, int>(dse.SystemFields.Id, Convert.ToInt32(Math.Round(dseAssembly.Amount)));

                             var dseAssemblyDataForDelete = ReadData(dseId, m_writeToLog, reportParameters);
                          //   Clipboard.SetText(string.Join("\r\n", dseAssemblyDataForDelete.Select(t => t.NomObj + " " + t.Amount)));
                          //   MessageBox.Show("1 " + string.Join("\r\n", dseAssemblyDataForDelete.Select(t => t.NomObj + " " + t.Amount)));
                             reportData = reportData.Where(t => dseAssemblyDataForDelete.Count(d=>d.NomObj.SystemFields.Id == t.NomObj.SystemFields.Id && d.Amount == t.Amount) == 0).ToList();

                             var dseObjectForRemove = new List<TFDDocument>();

                             foreach (var doc in reportData)
                             {
                                 var findedStructureObjectDSE = dseAssemblyDataForDelete.Where(t => t.Equals(doc)).FirstOrDefault();
                                 if (findedStructureObjectDSE != null)
                                 { 
                                     if (findedStructureObjectDSE.Amount == doc.Amount)
                                     {
                                         dseObjectForRemove.Add(doc);
                                     }
                                     else
                                     {
                                         if (doc.Amount > findedStructureObjectDSE.Amount)
                                         {
                                             doc.Amount = doc.Amount - findedStructureObjectDSE.Amount;
                                             doc.Position = doc.Position.Where(t => !findedStructureObjectDSE.Position.Contains(t)).ToList();
                                             doc.Remarks = doc.Remarks.Where(t => !findedStructureObjectDSE.Remarks.Contains(t)).ToList();
                                             doc.Zone = doc.Zone.Where(t => !findedStructureObjectDSE.Zone.Contains(t)).ToList();
                                             doc.MeasureUnit = doc.MeasureUnit.Where(t => !findedStructureObjectDSE.MeasureUnit.Contains(t)).ToList();
                                         }
                                         else
                                             MessageBox.Show("Обратитесь в Отдел 911. Некорректный подсчет объектов ДСЕ.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                     }
                                 }
                             }

                             reportData = reportData.Where(t => !dseObjectForRemove.Contains(t)).ToList();
                         }
                     }*/

                }
            }
            else
            {
                //Чтение данных для отчета    
                foreach (var id in reportParameters.listObjectCountId)
                {
                    var treeDevice = report.ReadData(id.Key, context, reportParameters, m_writeToLog);

                    foreach (var obj in treeDevice)
                    {
                        obj.Amount = obj.Amount * id.Value;
                    }
                    reportData.AddRange(treeDevice);
                }
            }

            if (reportData.Count == 1 || reportData.Count == 0)
            {
                //MessageBox.Show(
                //    "Объект выбранный для формирования отчета не имеет состава или выборка с дополнительными условиями возвращает нулевой состав.",
                //    "Внимание!",
                //    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //this.m_form.Close();
                //return;
            }
            if (getDocuments != null)
            {
                m_form.setStage(IndicationForm.Stages.DataProcessing);
                m_writeToLog("=== Обработка данных ===");

                sortedReportData = SumDublicates(SortData(reportData));
                getDocuments(sortedReportData, m_form, m_writeToLog);
            }
            else
            {
                try
                {
                    if (!reportParameters.LoadFolderName.Contains("Выгрузка"))
                    {
                        xls.Application.DisplayAlerts = false;
                        xls.Open(context.ReportFilePath);
                    }

                    m_form.setStage(IndicationForm.Stages.DataProcessing);
                    m_writeToLog("=== Обработка данных ===");

                    if (!reportParameters.OutputTechProcesses || reportParameters.GroupByTypes)
                    {
                        if (reportParameters.WithOutKooperationDSE)
                            sortedReportData = reportData.OrderBy(t => GetSortGroupByClass(t.NomObj)).ThenBy(t => t.Denotation).ThenBy(t => t.Naimenovanie).ToList();
                        else
                        {
                            if (reportParameters.MeasureUnit)
                                sortedReportData = SumDublicatesWithMeasureUnit(SortData(reportData));
                            else
                                sortedReportData = SumDublicates(SortData(reportData));
                        }
                    }
                    else
                    {
                        sortedReportData = new List<TFDDocument>();
                        sortedReportData.AddRange(reportData);
                    }
                    m_form.setStage(IndicationForm.Stages.ReportGenerating);
                    m_writeToLog("=== Формирование отчета ===");

                    //Формирование отчета
                    MakeSelectionByTypeReport(xls, sortedReportData, reportParameters, m_writeToLog);

                    m_form.progressBarParam(2);

                    m_form.setStage(IndicationForm.Stages.Done);
                    m_writeToLog("=== Завершение работы ===");

                    xls.AutoWidth();
                    if (!reportParameters.LoadFolderName.Contains("Выгрузка"))
                        xls.Save();
                    else
                    {
                        var listRootObj = reportData.Where(t => t.Level == 0).ToList();
                        var fileName = "Выборка по типам.xls";
                        if (listRootObj.Count == 1)
                        {
                            fileName = (listRootObj[0].Naimenovanie + " " + listRootObj[0].Denotation).Length < 100 ?
                                listRootObj[0].Naimenovanie + " " + listRootObj[0].Denotation + ".xlsx" :
                                listRootObj[0].Denotation + ".xlsx";
                        }
                        try
                        {
                            xls.SaveAs(reportParameters.LoadFolderName + fileName);
                            m_form.Dispose();
                        }
                        catch
                        {
                            MessageBox.Show(reportParameters.LoadFolderName + fileName);
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "Exception!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    isCatch = true;
                }
                finally
                {
                    // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                    if (isCatch)
                    {
                        xls.Quit(false);

                    }
                    else
                    {
                        xls.Quit(true);
                        // Открытие файла
                        if (!reportParameters.LoadFolderName.Contains("Выгрузка"))
                            System.Diagnostics.Process.Start(context.ReportFilePath);
                    }
                }
            }

            m_form.Dispose();
        }


        public List<TFDDocument> ReadData(KeyValuePair<int, int> id, LogDelegate logDelegate, ReportParameters reportParameters, List<ReferenceObject> dseObjects)
        {
            List<TFDDocument> objectsWithCount = new List<TFDDocument>();

            ReferenceInfo referenceNomenclatureInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Nomenclature.ReferenceId);
            Reference referenceNomenclature = referenceNomenclatureInfo.CreateReference();
            TFlex.DOCs.Model.Search.Filter filter = new TFlex.DOCs.Model.Search.Filter(referenceNomenclatureInfo);
            filter.Terms.AddTerm(referenceNomenclature.ParameterGroup[SystemParameterType.ObjectId],
                ComparisonOperator.Equal, id.Key);
            var nomObject = (NomenclatureObject)referenceNomenclature.Find(filter).First();
            List<TFDDocument> objects = new List<TFDDocument>();
            var nomReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            var reference = nomReferenceInfo.CreateReference();
            var hierarchyLinks = nomObject.Children.RecursiveLoadHierarchyLinks();
            var hierarchy = new Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>>();

            var hierarchyObject = new Dictionary<ReferenceObject, Dictionary<Guid, ComplexHierarchyLink>>();
            var objectsDictionary = new Dictionary<int, NomenclatureObject>();
            var objectsWithHierarchyDictionary = new Dictionary<ComplexHierarchyLink, NomenclatureObject>();
            objectsDictionary[nomObject.SystemFields.Id] = nomObject;

            //ДСЕ, которые присутствуют в в объекте для отчета
            var dseContainsInHierarchy = dseObjects.Where(t => hierarchyLinks.Any(d => t.SystemFields.Id == d.ChildObjectId)).ToList();
            var listDSEObject = new List<ComplexHierarchyLink>();

            if (dseContainsInHierarchy != null && dseContainsInHierarchy.Count > 0)
            {
                //Набираем все найденные ДСЕ иерархии в список
                foreach (var dse in dseContainsInHierarchy)
                {
                    listDSEObject.AddRange(dse.Children.RecursiveLoadHierarchyLinks());
                }
            }

            foreach (var hierarchyLink in hierarchyLinks.Where(t => !listDSEObject.Contains(t)))
            {
                var nhl = (NomenclatureHierarchyLink)hierarchyLink;
                objectsDictionary[hierarchyLink.ChildObjectId] = (NomenclatureObject)hierarchyLink.ChildObject;
                objectsWithHierarchyDictionary[nhl] = (NomenclatureObject)hierarchyLink.ChildObject;
                Dictionary<Guid, ComplexHierarchyLink> parentHierarchy = null;
                if (!hierarchy.TryGetValue(hierarchyLink.ParentObjectId, out parentHierarchy))
                {
                    parentHierarchy = new Dictionary<Guid, ComplexHierarchyLink>();
                    hierarchy[hierarchyLink.ParentObjectId] = parentHierarchy;
                    hierarchyObject[hierarchyLink.ParentObject] = parentHierarchy;
                }
                parentHierarchy[hierarchyLink.Guid] = hierarchyLink;
                logDelegate("Добавлен объект: " + nhl);
            }

            var documents = new List<TFDDocument>();

            GetReportDataWithCount(new TFDDocument(nomObject, reportParameters), 1, hierarchyObject, reportParameters, documents);

            // Clipboard.SetText(string.Join("\r\n", documents.OrderBy(t=> t.Naimenovanie + " " + t.Denotation).Select(t=>t.Naimenovanie + " " + t.Denotation + " - " +  t.Amount)));
            // MessageBox.Show("333");

            var objectTypes = GetSelectedObjectTypes(reportParameters);

            foreach (var obj in documents.Where(t => objectTypes.Any(t1 => t.NomObj.Class.IsInherit(t1)))
                .OrderBy(t => t.NomObj[Guids.Nomenclature.Fields.Denotation].GetString()).ThenBy(t => t.NomObj[Guids.Nomenclature.Fields.Name].GetString()))
            {
                obj.Amount = obj.Amount * id.Value;
                objects.Add(obj);
            }

            SetNomenclatureNumber(objects);

            foreach (var obj in objects)
            {
                if (obj.NomObj.SystemFields.Id == id.Key)
                    obj.Level = 0;
                else
                    obj.Level = 10;
            }

            return objects;
        }

        public List<Guid> GetSelectedObjectTypes(ReportParameters reportParameters)
        {
            var objectTypes = new List<Guid>();

            if (reportParameters.Assembly)
                objectTypes.Add(Guids.Nomenclature.Types.Assembly);

            if (reportParameters.Detail)
                objectTypes.Add(Guids.Nomenclature.Types.Detail);

            if (reportParameters.Documentation)
                objectTypes.Add(Guids.Nomenclature.Types.Document);

            if (reportParameters.Materials)
                objectTypes.Add(Guids.Nomenclature.Types.Material);

            if (reportParameters.Other)
                objectTypes.Add(Guids.Nomenclature.Types.OtherItem);

            if (reportParameters.Standart)
                objectTypes.Add(Guids.Nomenclature.Types.StandartItem);

            if (reportParameters.ComponentProgram)
                objectTypes.Add(Guids.Nomenclature.Types.ComponentProgram);

            if (reportParameters.Complex)
                objectTypes.Add(Guids.Nomenclature.Types.Complex);

            if (reportParameters.Complement)
                objectTypes.Add(Guids.Nomenclature.Types.Komplekt);

            if (reportParameters.ComplexProgram)
                objectTypes.Add(Guids.Nomenclature.Types.ComplexProgram);

            return objectTypes;
        }

        public void GetReportDataWithCount(TFDDocument parent, double count,
               Dictionary<ReferenceObject, Dictionary<Guid, ComplexHierarchyLink>> hierarchy,
               ReportParameters reportParameters, List<TFDDocument> documents)
        {
            var doc = documents.Where(t => t.Equals(parent)).FirstOrDefault();

            if (doc != null)
            {
                doc.Amount += count;
            }
            else
            {
                parent.Amount = count;
                documents.Add(parent);
            }

            if (hierarchy.ContainsKey(parent.NomObj))
            {
                foreach (var link in hierarchy[parent.NomObj].Values)
                {
                    var remark = link[Guids.NomHierarchyLink.Fields.Comment].Value.ToString();
                    var countNHL = Convert.ToDouble(link[Guids.NomHierarchyLink.Fields.Count].Value);

                    if (countNHL != Math.Truncate(countNHL) && link.ChildObject.Class.IsInherit(Guids.Nomenclature.Types.OtherItem))
                    {
                        countNHL = CountSelectionRezistor((NomenclatureObject)link.ChildObject, remark);
                    }

                    var linkCount = countNHL * count;
                    GetReportDataWithCount(new TFDDocument((NomenclatureObject)link.ChildObject, link, reportParameters), linkCount, hierarchy, reportParameters, documents);
                }
            }
        }

        private static int CountSelectionRezistor(NomenclatureObject nomObj, string remark)
        {
            var remarkWithOutLetter = remark.Replace("*", "");
            int errorId = 0;
            Regex regexDigitalOne = new Regex(@"(C|С|R|W)\d{1,}");
            Regex regexDigital = new Regex(@"(?<=(C|С|R|W))\d{1,}");
            Regex regexDigitalTwo = new Regex(@"(C|С|R|W)\d{1,}(-|(\s{1,})?\.{3}(\s{1,})?)(C|С|R|W)\d{1,}");
            int countRezist = 0;
            var matchrez = string.Empty;
            try
            {
                errorId = 1;
                foreach (Match match in regexDigitalTwo.Matches(remarkWithOutLetter))
                {
                    errorId = 2;
                    matchrez = regexDigital.Matches(match.Value.ToString())[0].Value + " " + regexDigital.Matches(match.Value.ToString())[1].Value;
                    countRezist += Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[1].Value) - Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[0].Value) + 1;
                    errorId = 3;
                    remarkWithOutLetter = remarkWithOutLetter.Replace(match.Value.ToString(), "");
                    errorId = 4;
                }
                errorId = 5;
                countRezist += regexDigitalOne.Matches(remarkWithOutLetter).Count;
                errorId = 6;
            }
            catch
            {
                MessageBox.Show("Обратитесь в отдел 911. Сформированный отчет будет некорректен. Код ошибки 7. Не могу посчитать количество резистора\r\n" + nomObj + "\r\nОшибка в примечании: " + remark + "\r\nКод ошибки " + errorId + "\r\n" + matchrez, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            return countRezist;
        }

        public List<ReferenceObject> GetDSEKooperation()
        {
            ReferenceInfo referenceDSEKooperationInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.DSEKooperation.ReferenceId);
            Reference referenceDSEKooperation = referenceDSEKooperationInfo.CreateReference();

            referenceDSEKooperation.Objects.Load();
            

            ReferenceInfo referenceNomenclatureInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Nomenclature.ReferenceId);
            Reference referenceNomenclature = referenceNomenclatureInfo.CreateReference();

            TFlex.DOCs.Model.Search.Filter filter = new TFlex.DOCs.Model.Search.Filter(referenceNomenclatureInfo);
            filter.Terms.AddTerm(referenceNomenclature.ParameterGroup[SystemParameterType.Guid],
                ComparisonOperator.IsOneOf,
                referenceDSEKooperation.Objects.GetObjects().Select(t => t[Guids.DSEKooperation.GuidParameter].GetGuid()));

            var nomKooperationDSE = referenceNomenclature.Find(filter);

            return nomKooperationDSE;
        }

        public void SetNomenclatureNumber(List<TFDDocument> documents)
        {
            ReferenceInfo referenceNomenclatureNumberInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.NomencaltureNumber.ReferenceID);
            Reference referenceNomenclatureNumber = referenceNomenclatureNumberInfo.CreateReference();

            foreach (var doc in documents)
            {
                TFlex.DOCs.Model.Search.Filter filter = new TFlex.DOCs.Model.Search.Filter(referenceNomenclatureNumberInfo);
                filter.Terms.AddTerm("[" + Guids.NomencaltureNumber.LinkToNomenclatureObject + "]", ComparisonOperator.Equal, doc.NomObj);

                var number = referenceNomenclatureNumber.Find(filter).FirstOrDefault();
                if (number != null)
                    doc.NomenclatureNumber = number[Guids.NomencaltureNumber.Number].GetString();
            }
        }

        // Чтение данных для отчета
        public List<TFDDocument> ReadData(int documentID, IReportGenerationContext context, ReportParameters reportParameters, LogDelegate logDelegate)
        {
            // соединяемся с БД T-FLEX DOCs 2010
            int errorId = 0;

            SqlConnection conn = null;

            errorId = 1;
            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            errorId = 2;
            try
            {
                //Создаем ссылку на справочник
                ReferenceInfo info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
                Reference reference = info.CreateReference();
                var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
                sqlConStringBuilder.DataSource = nameSqlServer;
                errorId = 3;
                sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
                errorId = 4;
                sqlConStringBuilder.Password = "reportUser";
                errorId = 5;
                sqlConStringBuilder.UserID = "reportUser";
                errorId = 6;
                conn = new SqlConnection(sqlConStringBuilder.ToString());
                errorId = 7;
                conn.Open();
                errorId = 0;
            }
            catch (Exception ex)
            {
                Clipboard.SetText(sqlConStringBuilder.ToString());
                MessageBox.Show(errorId + "\r\n" + ex.ToString());
            }

            List<TFDDocument> objects = new List<TFDDocument>();

            //Формирование списка классов для отчета (Вывод техпроцессов)
            string listReportTPClass = string.Empty;
            if (reportParameters.DetailTP)
            {
                listReportTPClass += TFDClass.Detal.ToString();
            }

            if (reportParameters.AssemblyTP)
            {
                if (listReportTPClass != string.Empty)
                    listReportTPClass += "," + TFDClass.Assembly.ToString();
                else
                    listReportTPClass += TFDClass.Assembly.ToString();
            }

            if (reportParameters.ComplementTP)
            {
                if (listReportTPClass != string.Empty)
                    listReportTPClass += "," + TFDClass.Komplekt.ToString();
                else
                    listReportTPClass += TFDClass.Komplekt.ToString();
            }


            if (reportParameters.ComplexTP)
            {
                if (listReportTPClass != string.Empty)
                    listReportTPClass += "," + TFDClass.Complex.ToString();
                else
                    listReportTPClass += TFDClass.Complex.ToString();
            }


            if (reportParameters.StandartItemsTP)
            {
                if (listReportTPClass != string.Empty)
                    listReportTPClass += "," + TFDClass.StandartItem.ToString();
                else
                    listReportTPClass += TFDClass.StandartItem.ToString();
            }

            // Формирование списка объектов, которые исключаются из состава объекта для формирования отчета
            string listDelObj = string.Empty;

            if (reportParameters.listDeleteObjectId.Count > 1)
            {
                for (int i = 0; i < reportParameters.listDeleteObjectId.Count - 1; i++)
                {
                    listDelObj += reportParameters.listDeleteObjectId[i] + ",";
                }
            }
            listDelObj += reportParameters.listDeleteObjectId[reportParameters.listDeleteObjectId.Count - 1].ToString();

            if (!reportParameters.ExpansionNodes)
            {
                #region запрос для отчета без разузловки
                string getDocTreeCommandText = String.Format(@"
                              SELECT DISTINCT  nh.s_ObjectID,
                                     1,                    -- level
		                             nh.s_ParentID,
		                             n.s_ClassID,
		                             nh.Position,
		                             nh.Zone,
		                             n.Format,
		                             n.Denotation,
		                             n.Name,
		                             nh.Amount,
		                             nh.Remarks,
                                     n.letterEx,
                                     bdp.CodeRCP,
                                     cp.FirstUse,
                                     n.s_CreationDate,
                                     nn.Number,
                                     nh.MeasureUnit,
									 n.s_AuthorID,
									 n.s_EditorID
                            FROM NomenclatureHierarchy nh JOIN Nomenclature  n ON n.s_ObjectID = nh.s_ObjectID    
                            LEFT JOIN BuyDeviceParams bdp ON (nh.s_ObjectID = bdp.s_ObjectID) AND bdp.s_ActualVersion = 1 AND bdp.s_Deleted = 0
                            LEFT JOIN ConstructorParams cp ON (n.s_ObjectID = cp.s_ObjectID) AND cp.s_ActualVersion = 1 AND cp.s_Deleted = 0
                            LEFT JOIN NomenclatureObjectLink nol ON (n.s_ObjectID = nol.SlaveID) 
                            LEFT JOIN NomenclatureNumbers nn ON (nn.s_ObjectID = nol.MasterID) AND nn.s_ActualVersion = 1 AND nn.s_Deleted = 0
                            WHERE n.s_ActualVersion = 1 AND n.s_Deleted = 0 AND
                                  nh.s_ActualVersion = 1 AND nh.s_Deleted = 0 AND
                                  nh.s_ParentID = {0} AND
                                  n.s_ObjectID NOT IN ({1})
                            ---- AddWhereDenotation
                            --AddWhereName
                            UNION ALL
                            SELECT n.s_ObjectID,
                                   0,            --level
	                               0,
	                               n.s_ClassID,
	                               0,
	                               '',
	                               '',
	                               n.Denotation,
	                               n.Name,
	                               1,
	                               '',
                                   n.letterEx,
                                   null,
                                   '',
                                   n.s_CreationDate,
                                   '',
                                   '',
							       n.s_AuthorID,
								   n.s_EditorID
                            FROM Nomenclature n
                            WHERE n.s_ActualVersion = 1 AND s_Deleted = 0 AND 
                                  n.s_ObjectID = {0}", documentID, listDelObj);



                int indx = 0;

                if (reportParameters.AddLikeName.Count > 0)
                {
                    indx = getDocTreeCommandText.IndexOf("----") - 1;
                    getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "AND (");
                }

                for (int i = 0; i < reportParameters.AddLikeName.Count; i++)
                {
                    indx = getDocTreeCommandText.IndexOf("----") - 1;


                    if (reportParameters.AddLikeDenotation[i] != null && reportParameters.AddLikeDenotation[i] != string.Empty)
                    {
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                              string.Format(
                                                                                  @" (tv.Denotation Like N'%{0}%' ",
                                                                                  reportParameters.AddLikeDenotation[i]));


                        if (reportParameters.AddLikeName[i] != null && reportParameters.AddLikeName[i] != string.Empty)
                        {
                            indx = getDocTreeCommandText.IndexOf("----") - 1;
                            getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                                  string.Format(
                                                                                      @" AND tv.[Name] Like N'%{0}%') ",
                                                                                      reportParameters.AddLikeName[i]));
                        }
                        else
                        {
                            indx = getDocTreeCommandText.IndexOf("----") - 1;
                            getDocTreeCommandText = getDocTreeCommandText.Insert(indx, ")");
                        }
                    }
                    else
                    {
                        if (reportParameters.AddLikeName[i] != null && reportParameters.AddLikeName[i] != string.Empty)
                        {
                            indx = getDocTreeCommandText.IndexOf("----") - 1;
                            getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                                  string.Format(
                                                                                      @" tv.[Name] Like N'%{0}%' ",
                                                                                      reportParameters.AddLikeName[i]));
                        }
                    }
                    if (i < reportParameters.AddLikeName.Count - 1)
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "OR");
                    }
                }

                if (reportParameters.AddLikeName.Count > 0)
                {
                    indx = getDocTreeCommandText.IndexOf("----") - 1;
                    getDocTreeCommandText = getDocTreeCommandText.Insert(indx, ")");
                }



                var docTreeCommand = new SqlCommand(getDocTreeCommandText, conn);
                docTreeCommand.CommandTimeout = 0;
                SqlDataReader reader = docTreeCommand.ExecuteReader();

                TFDDocument doc;

                while (reader.Read())
                {
                    doc = new TFDDocument();
                    doc.ObjectID = GetSqlInt32(reader, 0);
                    doc.Level = GetSqlInt32(reader, 1);
                    doc.ParentID = GetSqlInt32(reader, 2);
                    doc.Class = GetSqlInt32(reader, 3);

                    var position = GetSqlInt32(reader, 4).ToString().Trim();
                    if (position != string.Empty)
                        doc.Position.Add(position);

                    var zone = GetSqlString(reader, 5).Trim();
                    if (zone != string.Empty)
                        doc.Zone.Add(zone);

                    doc.Format = GetSqlString(reader, 6);
                    doc.Denotation = GetSqlString(reader, 7);
                    doc.Naimenovanie = GetSqlString(reader, 8);
                    doc.Amount = GetSqlDouble(reader, 9);

                    var remarks = GetSqlString(reader, 10).Trim();
                    if (remarks != string.Empty)
                        doc.Remarks.Add(remarks);

                    doc.CodeRCP = "'" + GetSqlString(reader, 12);
                    doc.FirstUse = "'" + GetSqlString(reader, 13);
                    doc.CreateDate = reader.GetDateTime(14).ToShortDateString();
                    doc.NomenclatureNumber = GetSqlString(reader, 15);

                    var measureUnit = GetSqlString(reader, 16).Trim();
                    if (measureUnit != string.Empty)
                        doc.MeasureUnit.Add(measureUnit);

                    if (reportParameters.AuthorName)
                    {
                        doc.AuthorName = GetShortNameUserByID(GetSqlInt32(reader, 17));
                    }

                    if (reportParameters.LastEditorName)
                    {
                        doc.LastEditorName = GetShortNameUserByID(GetSqlInt32(reader, 18));
                    }

                    doc.Letter = GetSqlString(reader, 11);
                    if ((reportParameters.Assembly) && (TFDClass.Assembly == doc.Class) || (reportParameters.Complement) && (TFDClass.Komplekt == doc.Class) ||
                        (reportParameters.Complex) && (TFDClass.Complex == doc.Class) || (reportParameters.ComplexProgram) && (TFDClass.ComplexProgram == doc.Class) ||
                        (reportParameters.ComponentProgram) && (TFDClass.ComponentProgram == doc.Class) || (reportParameters.Detail) && (TFDClass.Detal == doc.Class) ||
                        (reportParameters.Documentation) && (TFDClass.Document == doc.Class) || (reportParameters.Materials) && (TFDClass.Material == doc.Class) ||
                        (reportParameters.Other) && (TFDClass.OtherItem == doc.Class) || (reportParameters.Standart) && (TFDClass.StandartItem == doc.Class) || (doc.Level == 0))
                    {
                        objects.Add(doc);
                        logDelegate(String.Format("Добавлен объект: {0} {1}", doc.Denotation, doc.Naimenovanie));
                    }
                }
                m_form.progressBarParam(2);
                reader.Close();

                #endregion
            }
            else
            {
                #region запрос для отчета с разузловкой
                //запрос для отчета с разузловкой

                string getDocTreeCommandTextOnePart = String.Format(@"
                                
                                declare @docid INT
                                DECLARE @level int
                                declare @insertcount int
                                set @docid={0}
                                set @insertcount = 0
                                SET @level=0

                                IF OBJECT_ID('tempdb..#TmpVspec')is NULL
                                CREATE TABLE #TmpVspec (s_ObjectID INT,
                                                        [level] INT,
                                                        s_ParentID INT,
                                                        s_ClassID INT,
                                                        Position INT,
                                                        Zone NVARCHAR(20),
                                                        Format NVARCHAR(20),
                                                        Denotation NVARCHAR(255),
                                                        Name NVARCHAR(255),
                                                        Amount FLOAT,
                                                        Remarks NVARCHAR(MAX),
                                                        letterEx NVARCHAR(50),
                                                        TotalCount FLOAT,
                                                        CreateDate DATETIME,
                                                        MeasureUnit NVARCHAR(255),
                                                        AuthorName INT,
                                                        LastEditorName INT)
                                ELSE DELETE FROM #TmpVspec

                                INSERT INTO #TmpVspec
                                SELECT n.s_ObjectID,0,0,n.s_ClassID,0,'','',n.Denotation,n.Name,0,'',n.letterEx,1,n.s_CreationDate, '', n.s_AuthorID, n.s_EditorID
                                FROM Nomenclature n
                                WHERE n.s_ObjectID = @docid
                                      AND n.s_Deleted = 0 
                                      AND n.s_ActualVersion = 1 

                                WHILE 1=1
                                BEGIN

                                  INSERT INTO #TmpVspec 
                                  SELECT nh.s_ObjectID,
                                         @level+1,
                                         nh.s_ParentID,
                                         n.s_ClassID,
                                         nh.Position,
                                         nh.Zone,
                                         n.Format,
                                         n.Denotation,
                                         n.Name,
                                         nh.Amount,
                                         nh.Remarks,
                                         n.letterEx,
                                         #TmpVspec.TotalCount*nh.Amount,
                                         n.s_CreationDate,
                                         nh.MeasureUnit,
                                         n.s_AuthorID,
									     n.s_EditorID
                                  FROM NomenclatureHierarchy nh INNER JOIN #TmpVspec ON nh.s_ParentID=#TmpVspec.s_ObjectID 
                                                                           
                                       INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                                  WHERE
                                      #TmpVspec.[level]=@level
                                      AND nh.s_ActualVersion = 1
                                      AND nh.s_Deleted = 0
                                      AND n.s_ActualVersion = 1
                                      AND n.s_Deleted = 0
                                      AND n.s_ObjectID NOT IN ({1})
                                  
                                  SET @insertcount = @@ROWCOUNT 
                                  SET @level = @level + 1 
   
                                  IF @insertcount = 0 
                                  GOTO end1

                                END
                                end1:

                                  
                                          ", documentID, listDelObj);

                string getDocTreeCommandTextSecondPart = String.Format(@"
                                SELECT tv.s_ObjectID,
                                       tv.level, 
                                       tv.s_ParentID, 
                                       tv.s_ClassID, 
                                       tv.Position, 
                                       tv.Zone, 
                                       tv.Format, 
                                       tv.Denotation, 
                                       tv.Name, 
                                       SUM(tv.TotalCount)as Summa, 
                                       tv.Remarks, 
                                       tv.letterEx,
                                       tv.Amount,
                                       bdp.CodeRCP,
                                       cp.FirstUse,
                                       tv.CreateDate,
                                       nn.Number,
                                       tv.MeasureUnit,
                                       tv.AuthorName,
                                       tv.LastEditorName
                          
                                FROM #TmpVspec tv LEFT  JOIN BuyDeviceParams bdp ON (tv.s_ObjectID = bdp.s_ObjectID) AND bdp.s_ActualVersion = 1 AND bdp.s_Deleted = 0
                                      LEFT JOIN ConstructorParams cp ON (tv.s_ObjectID = cp.s_ObjectID) AND cp.s_ActualVersion = 1 AND cp.s_Deleted = 0
                                      LEFT JOIN NomenclatureObjectLink nol ON (tv.s_ObjectID = nol.SlaveID) 
                                      LEFT JOIN NomenclatureNumbers nn ON (nn.s_ObjectID = nol.MasterID) AND nn.s_ActualVersion = 1 AND nn.s_Deleted = 0
                                ----AddWhereDenotation
                                --AddWhereName
         
                                GROUP BY tv.s_ObjectID,  
                                         tv.s_ParentID, 
                                         tv.s_ClassID, 
                                         tv.Position, 
                                         tv.Zone, 
                                         tv.Format, 
                                         tv.Denotation, 
                                         tv.Name, 
                                         tv.Remarks, 
                                         tv.letterEx,
                                         tv.level,
                                         tv.Amount,
                                         bdp.CodeRCP,
                                         cp.FirstUse,
                                         tv.CreateDate,
                                         nn.Number,
                                         tv.MeasureUnit,
                                         tv.AuthorName,
                                         tv.LastEditorName");

                if (reportParameters.OutputTechProcesses)
                {
                    getDocTreeCommandTextSecondPart = String.Format(@"
                                       SELECT DISTINCT
                                               n.[s_ObjectID]
                                              ,n.[Name]
                                              ,n.[Denotation]
                                              ,tp.TPName
                                              ,NULL as SumPieceTime
                                              ,NULL as SumPrepTime
                                              ,NULL as primechanie
                                              ,NULL as PieceTime
                                              ,n.s_ClassID
                                              ,'' as OperTxt
                                              ,tp.s_ObjectID
                                              ,''
                                              , n.s_CreationDate
                                              ,''
                                              ,''
                                              ,n.s_AuthorID
									          ,n.s_EditorID
                                        FROM  #TmpVspec tv LEFT JOIN [TFlexDOCs].[dbo].[Nomenclature] n ON (tv.s_ObjectID = n.s_ObjectID) 
                                               LEFT JOIN TPLinksNomenclatureDSE tpn ON (n.s_ObjectID = tpn.SlaveID)
                                               LEFT JOIN TechnologicalProcesses tp ON (tp.s_ObjectID = tpn.MasterID)    
                                        WHERE n.s_ActualVersion = 1 AND n.s_Deleted = 0  AND (tv.[level] = 0  OR n.s_ClassID in ({0})) AND (tp.TPName IS NULL OR DeletedChangelistID <> 0)

                                        UNION 

                                        SELECT DISTINCT
                                               n.[s_ObjectID]
                                              ,n.[Name]
                                              ,n.[Denotation]
                                              ,tp.TPName
                                              ,tpd.SumPieceTime
                                              ,tpd.SumPrepTime
                                              ,tps.primechanie
                                              ,tpo.PieceTime
                                              ,n.s_ClassID
                                              ,tpo.OperTxt  
                                              ,tp.s_ObjectID
                                              ,tpo.SerialNumber
                                              ,n.s_CreationDate
                                              ,''
                                              ,''
                                              ,n.s_AuthorID
									          ,n.s_EditorID
                                        FROM  #TmpVspec tv LEFT JOIN [TFlexDOCs].[dbo].[Nomenclature] n ON (tv.s_ObjectID = n.s_ObjectID) 
                                              LEFT JOIN TPLinksNomenclatureDSE tpn ON (n.s_ObjectID = tpn.SlaveID)
                                              LEFT JOIN TechnologicalProcesses tp ON (tp.s_ObjectID = tpn.MasterID)
                                              LEFT JOIN TechnologicalProcessData tpd ON (tpd.s_ObjectID = tp.s_ObjectID)
                                              LEFT JOIN glob_TP_Signs tps ON (tp.s_ObjectID = tps.s_ObjectID)
                                              LEFT JOIN TPOperations tpo ON (tp.s_ObjectID = tpo.s_MasterID)
                                        WHERE n.s_ActualVersion = 1 AND n.s_Deleted = 0  AND (tv.[level] = 0  OR n.s_ClassID in ({0})) AND DeletedChangelistID = 0 AND
                                              tp.s_ActualVersion = 1  AND tp.s_Deleted = 0  AND 
                                              tpd.s_ActualVersion = 1 AND tpd.s_Deleted = 0  AND
                                              tps.s_ActualVersion = 1 AND tps.s_Deleted = 0 AND
                                              tpo.s_ActualVersion = 1 AND tpo.s_Deleted = 0 AND DeletedChangelistID = 0  
                                        ORDER BY
                                              n.s_ClassID
                                             ,n.[Denotation]
                                             ,tp.s_ObjectID", listReportTPClass);
                }

                string getDocTreeCommandText = getDocTreeCommandTextOnePart + getDocTreeCommandTextSecondPart;

                int indx = 0;
                if (reportParameters.AddLikeName.Count > 0)
                {
                    indx = getDocTreeCommandText.IndexOf("----") - 1;
                    getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "WHERE");
                }

                for (int i = 0; i < reportParameters.AddLikeName.Count; i++)
                {
                    indx = getDocTreeCommandText.IndexOf("----") - 1;


                    if (reportParameters.AddLikeDenotation[i] != null && reportParameters.AddLikeDenotation[i] != string.Empty)
                    {
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                              string.Format(
                                                                                  @" (tv.Denotation Like N'%{0}%' ",
                                                                                  reportParameters.AddLikeDenotation[i]));


                        if (reportParameters.AddLikeName[i] != null && reportParameters.AddLikeName[i] != string.Empty)
                        {
                            indx = getDocTreeCommandText.IndexOf("----") - 1;
                            getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                                  string.Format(
                                                                                      @" AND tv.[Name] Like N'%{0}%') ",
                                                                                      reportParameters.AddLikeName[i]));
                        }
                        else
                        {
                            indx = getDocTreeCommandText.IndexOf("----") - 1;
                            getDocTreeCommandText = getDocTreeCommandText.Insert(indx, ")");
                        }
                    }
                    else
                    {
                        if (reportParameters.AddLikeName[i] != null && reportParameters.AddLikeName[i] != string.Empty)
                        {
                            indx = getDocTreeCommandText.IndexOf("----") - 1;
                            getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                                  string.Format(
                                                                                      @" tv.[Name] Like N'%{0}%' ",
                                                                                      reportParameters.AddLikeName[i]));
                        }
                    }
                    if (i < reportParameters.AddLikeName.Count - 1)
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "OR");
                    }
                }


                var docTreeCommand = new SqlCommand(getDocTreeCommandText, conn);
                docTreeCommand.CommandTimeout = 0;
                SqlDataReader reader = docTreeCommand.ExecuteReader();

                TFDDocument doc;

                while (reader.Read())
                {
                    if (!reportParameters.OutputTechProcesses)
                    {
                        doc = new TFDDocument();
                        doc.ObjectID = GetSqlInt32(reader, 0);
                        doc.Level = GetSqlInt32(reader, 1);
                        doc.ParentID = GetSqlInt32(reader, 2);
                        doc.Class = GetSqlInt32(reader, 3);

                        var position = GetSqlInt32(reader, 4).ToString().Trim();
                        if (position != string.Empty)
                            doc.Position.Add(position);

                        var zone = GetSqlString(reader, 5).Trim();
                        if (zone != string.Empty)
                            doc.Zone.Add(zone);

                        doc.Format = GetSqlString(reader, 6);
                        doc.Denotation = GetSqlString(reader, 7);
                        doc.Naimenovanie = GetSqlString(reader, 8);
                        doc.Amount = GetSqlDouble(reader, 9);

                        var remarks = GetSqlString(reader, 10).Trim();
                        if (remarks != string.Empty)
                            doc.Remarks.Add(remarks);

                        doc.Letter = GetSqlString(reader, 11);

                        doc.CodeRCP = "'" + GetSqlString(reader, 13);
                        doc.FirstUse = "'" + GetSqlString(reader, 14);
                        doc.CreateDate = reader.GetDateTime(15).ToShortDateString();
                        doc.NomenclatureNumber = GetSqlString(reader, 16);

                        var measureUnit = GetSqlString(reader, 17).Trim();
                        if (measureUnit != string.Empty)
                            doc.MeasureUnit.Add(measureUnit);

                        if (reportParameters.AuthorName)
                        {
                            doc.AuthorName = GetShortNameUserByID(GetSqlInt32(reader, 18));
                        }

                        if (reportParameters.LastEditorName)
                        {
                            doc.LastEditorName = GetShortNameUserByID(GetSqlInt32(reader, 19));
                        }

                        if (doc.Amount != Math.Truncate(doc.Amount) && (TFDClass.OtherItem == doc.Class) && (doc.Naimenovanie.Contains("Резистор") || doc.Naimenovanie.Contains("Конденсатор")))
                        {
                            double amountByRemark = CalculateAmountByRemark(string.Join(" ", doc.Remarks), doc.Naimenovanie);

                            if (GetSqlDouble(reader, 12) > amountByRemark)
                                amountByRemark = GetSqlDouble(reader, 12);

                            doc.Amount = amountByRemark * GetSqlDouble(reader, 9) / GetSqlDouble(reader, 12);
                        }
                        if ((reportParameters.Assembly) && (TFDClass.Assembly == doc.Class) || (reportParameters.Complement) && (TFDClass.Komplekt == doc.Class) ||
                             (reportParameters.Complex) && (TFDClass.Complex == doc.Class) || (reportParameters.ComplexProgram) && (TFDClass.ComplexProgram == doc.Class) ||
                             (reportParameters.ComponentProgram) && (TFDClass.ComponentProgram == doc.Class) || (reportParameters.Detail) && (TFDClass.Detal == doc.Class) ||
                             (reportParameters.Documentation) && (TFDClass.Document == doc.Class) || (reportParameters.Materials) && (TFDClass.Material == doc.Class) ||
                             (reportParameters.Other) && (TFDClass.OtherItem == doc.Class) || (reportParameters.Standart) && (TFDClass.StandartItem == doc.Class) || (doc.Level == 0))
                        {
                            objects.Add(doc);
                            logDelegate(String.Format("Добавлен объект: {0} {1}", doc.Denotation, doc.Naimenovanie));
                        }
                    }
                    else
                    {
                        var techProc = new TechnologicalProcess();
                        techProc.Name = GetSqlString(reader, 3);
                        techProc.SumPieceTime = GetSqlDouble(reader, 4);
                        techProc.SumPrepTim = GetSqlDouble(reader, 5);
                        techProc.Remarks = GetSqlString(reader, 6);
                        techProc.PieceTime = GetSqlDouble(reader, 7);
                        techProc.OperName = GetSqlString(reader, 9);
                        techProc.ObjectID = GetSqlInt32(reader, 10);

                        if (objects.Where(t => t.ObjectID == GetSqlInt32(reader, 0)).ToList().Count == 0)
                        {
                            doc = new TFDDocument();
                            doc.ObjectID = GetSqlInt32(reader, 0);

                            if (doc.ObjectID == documentID)
                            {
                                doc.Level = 0;
                            }
                            else
                            {
                                doc.Level = 10;
                            }
                            doc.Naimenovanie = GetSqlString(reader, 1);
                            doc.Denotation = GetSqlString(reader, 2);
                            doc.Class = GetSqlInt32(reader, 8);
                            doc.parenDevicetId = documentID;
                            if (techProc.OperName.Trim() != string.Empty)
                                doc.TechProcesses.Add(techProc);
                            objects.Add(doc);
                        }
                        else
                        {
                            foreach (var obj in objects)
                            {
                                if (obj.ObjectID == GetSqlInt32(reader, 0) && techProc.OperName != string.Empty)
                                {
                                    obj.TechProcesses.Add(techProc);
                                    break;
                                }
                            }
                        }
                    }
                }

                m_form.progressBarParam(2);
                reader.Close();
                #endregion
            }
            return objects;
        }

        private string GetShortNameUserByID(int userId)
        {
            //Загрузка объектов справочника Шаблоны типовых документов
            ReferenceInfo referenceUserInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.RefId);
            Reference referenceUsers = referenceUserInfo.CreateReference();

            var filter = new TFlex.DOCs.Model.Search.Filter(referenceUserInfo);
            filter.Terms.AddTerm(referenceUsers.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, userId);

            var userObject = referenceUsers.Find(filter).FirstOrDefault();

            if (userObject == null)
            {
                return string.Empty;
            }
            var name = userObject[Guids.ГруппыИПользователи.Поля.КороткоеИмя].GetString().Trim();

            return name != string.Empty ? name : userObject.ToString().Trim();

        }

        private int CalculateAmountByRemark(string remark, string name)
        {
            var remarkWithOutLetter = remark.Replace("*", "");
            int errorId = 0;
            Regex regexDigitalOne = new Regex(@"(C|С|R)\d{1,}");
            Regex regexDigital = new Regex(@"(?<=(C|С|R))\d{1,}");
            Regex regexDigitalTwo = new Regex(@"(C|С|R)\d{1,}(-|\.{3})(C|С|R)\d{1,}");
            int countRezist = 0;
            var matchrez = string.Empty;
            try
            {
                errorId = 1;
                foreach (Match match in regexDigitalTwo.Matches(remarkWithOutLetter))
                {
                    errorId = 2;
                    matchrez = regexDigital.Matches(match.Value.ToString())[0].Value + " " + regexDigital.Matches(match.Value.ToString())[1].Value;
                    countRezist += Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[1].Value) - Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[0].Value) + 1;
                    errorId = 3;
                    remarkWithOutLetter = remarkWithOutLetter.Replace(match.Value.ToString(), "");
                    errorId = 4;
                }
                errorId = 5;
                countRezist += regexDigitalOne.Matches(remarkWithOutLetter).Count;
                errorId = 6;
            }
            catch
            {
                MessageBox.Show("Не могу посчитать количество резистора\r\n" + name + "\r\nОшибка в примечании: " + remark + "\r\nКод ошибки " + errorId + "\r\n" + matchrez, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return countRezist;
        }

        public static string GetSqlString(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return String.Empty;

            return reader.GetString(field);
        }


        public static int GetSqlInt32(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0;

            return reader.GetInt32(field);
        }

        public static double GetSqlDouble(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0d;

            return reader.GetDouble(field);
        }

        // Запись в ячейку элемента заголовка таблицы
        public static int InsertHeader(bool isCheck, string header, int col, int row, Xls xls)
        {
            if (isCheck)
            {
                xls[col, row].Font.Name = "Calibri";
                xls[col, row].Font.Size = 11;
                xls[col, row].SetValue(header);
                xls[col, row].Font.Bold = true;
                col++;
            }
            return col;
        }

        // Запись заголовка раздела спецификации
        public static void InsertBomSectionHeader(int row, int col, string bomSectionHeader, Xls xls)
        {
            xls[1, row].Font.Name = "Calibri";
            xls[1, row].Font.Size = 11;
            xls[1, row].SetValue(bomSectionHeader);
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Underline = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, col - 1, 1].Merge();
            xls[1, row].Interior.Color = Color.LightGray;
        }

        // Запись заголовка наименования устройства
        public static void InsertDeviceNameHeader(int row, int col, string bomSectionHeader, Xls xls)
        {
            xls[1, row].Font.Name = "Calibri";
            xls[1, row].Font.Size = 11;
            xls[1, row].SetValue(bomSectionHeader);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, row, col - 1, 1].Merge();
            xls[1, row].Interior.Color = Color.Yellow;
        }

        public static void InsertHeader(int row, int col, string bomSectionHeader, Xls xls)
        {
            xls[1, row].Font.Name = "Calibri";
            xls[1, row].Font.Size = 11;
            xls[1, row].SetValue(bomSectionHeader);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, row, col - 1, 1].Merge();
        }






        // Запись в ячейку параметра объекта
        public static int InsertCell(bool isCheck, int col, int row, string value, Xls xls)
        {
            if (isCheck)
            {
                if (value != null && value.Trim() != string.Empty)
                {
                    xls[col, row].Font.Name = "Calibri";
                    xls[col, row].Font.Size = 11;
                    xls[col, row].SetValue(value.Trim());
                }
                col++;
            }
            return col;
        }

        // Запись в ячейку параметра объекта
        public static int InsertCell(int col, int row, string value, Xls xls)
        {

            xls[col, row].SetValue(value);
            col++;
            return col;
        }

        // Запись в ячейку параметра объекта
        public static int InsertCell(bool isCheck, int col, int row, int value, Xls xls)
        {
            if (isCheck)
            {
                if (value != 0)
                {
                    xls[col, row].Font.Name = "Calibri";
                    xls[col, row].Font.Size = 11;
                    xls[col, row].SetValue(value);
                }
                col++;
            }
            return col;
        }

        // Запись в ячейку параметра объекта
        public static int InsertCell(bool isCheck, int col, int row, double value, Xls xls)
        {
            if (isCheck)
            {
                if (value != 0)
                {
                    xls[col, row].Font.Name = "Calibri";
                    xls[col, row].Font.Size = 11;
                    xls[col, row].SetValue(value);
                }
                col++;
            }
            return col;
        }

        public static string text(string[] textArray)
        {
            string str = string.Empty;
            if (textArray.Count() > 1)
            {
                for (int i = 0; i < textArray.Count() - 1; i++)
                {
                    if (textArray[i].Trim() != string.Empty)
                        str += textArray[i] + '\n';
                }
                str += textArray[textArray.Count() - 1];
            }
            else str = textArray[textArray.Count() - 1];

            return str;
        }

        public static int countStr = 0;
        public static int countWrap = 0;
        public static int countPage = 0;
        public static List<string> HaveFiles;

        public static int InsertObject(ReportParameters reportParameters, TFDDocument doc, int row, bool type, int classId, Xls xls, LogDelegate logDelegate)
        {
            var errorId = 0;

            try
            {
                int col = 1;
                if (type && (doc.Class == classId))
                {
                    errorId = 1;
                    logDelegate(String.Format("Запись объекта: {0} {1}", doc.Denotation, doc.Naimenovanie));
                    col = InsertCell(reportParameters.Position, col, row, string.Join("; ", doc.Position.Where(t => t.Trim() != string.Empty).OrderBy(t => t.ToString())), xls);
                    col = InsertCell(reportParameters.Zone, col, row, string.Join("; ", doc.Zone.Where(t => t.Trim() != string.Empty).OrderBy(t => t.ToString())), xls);
                    xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    col = InsertCell(reportParameters.Format, col, row, doc.Format, xls);
                    string denotation = doc.Denotation;
                    string name = doc.Naimenovanie;
                    string remarks = string.Join("; ", doc.Remarks.Where(t => t.Trim() != string.Empty).OrderBy(t => t.ToString()));
                    errorId = 2;
                    //  if ((classId == TFDClass.OtherItem) && reportParameters.ExpansionNodes)

                    //     remarks = string.Empty;

                    if (reportParameters.Main5Col)
                    {
                        errorId = 3;
                        Font font = new Font("Calibri", 11f, GraphicsUnit.Document);
                        TextFormatter _textFormatter = new TextFormatter(font);

                        string[] wrapTextDenotation = _textFormatter.Wrap(doc.Denotation.Trim(), 16f);
                        denotation = text(wrapTextDenotation);

                        string[] wrapTextName = _textFormatter.Wrap(doc.Naimenovanie.Trim(), 53f);
                        name = text(wrapTextName);

                        string[] wrapTextRemarks = _textFormatter.Wrap(remarks.Trim(), 16f);
                        remarks = text(wrapTextRemarks);

                        if (wrapTextRemarks.Count() >= wrapTextName.Count())
                        {
                            countWrap = wrapTextRemarks.Count();
                            if (wrapTextDenotation.Count() >= wrapTextRemarks.Count())
                                countWrap = wrapTextDenotation.Count();
                        }
                        else
                        {
                            countWrap = wrapTextName.Count();
                            if (wrapTextDenotation.Count() >= wrapTextName.Count())
                                countWrap = wrapTextDenotation.Count();
                        }
                        countStr += countWrap;
                    }
                    errorId = 4;
                    col = InsertCell(reportParameters.Denotation, col, row, denotation, xls);
                    errorId = 5;
                    col = InsertCell(reportParameters.Name, col, row, name, xls);
                    errorId = 6;
                    col = InsertCell(reportParameters.Count, col, row, doc.Amount, xls);
                    errorId = 7;
                    col = InsertCell(reportParameters.Remarks, col, row, remarks, xls);
                    errorId = 8;
                    col = InsertCell(reportParameters.Letter, col, row, doc.Letter, xls);
                    errorId = 9;
                    col = InsertCell(reportParameters.Id, col, row, doc.ObjectID, xls);
                    errorId = 10;

                    col = InsertCell(reportParameters.CodeRcp, col, row, doc.CodeRCP, xls);
                    errorId = 11;
                    col = InsertCell(reportParameters.FirstUse, col, row, doc.FirstUse, xls);
                    errorId = 12;
                    col = InsertCell(reportParameters.CreateDate, col, row, doc.CreateDate, xls);
                    errorId = 13;
                    col = InsertCell(reportParameters.NomenclatureNumber, col, row, doc.NomenclatureNumber, xls);
                    errorId = 14;
                    col = InsertCell(reportParameters.MeasureUnit, col, row, string.Join("; ", doc.MeasureUnit.Where(t => t.Trim() != string.Empty).OrderBy(t => t.ToString())), xls);
                    errorId = 15;
                    col = InsertCell(reportParameters.AuthorName, col, row, doc.AuthorName, xls);
                    col = InsertCell(reportParameters.LastEditorName, col, row, doc.LastEditorName, xls);
                    if (HaveFiles != null)
                    {
                        col = InsertCell(true, col, row, HaveFiles.Contains(denotation) ? "+" : "-", xls);
                    }


                    row++;
                }
            }
            catch //()
            {
                //  MessageBox.Show("111 - " + errorId);
            }
            return row;
        }


        public static int InsertObjectTP(ReportParameters reportParameters, TFDDocument doc, int row, bool type, int classId, Xls xls, LogDelegate logDelegate)
        {

            int col = 1;
            if (type && (doc.Class == classId))
            {

                logDelegate(String.Format("Запись объекта: {0} {1}", doc.Denotation, doc.Naimenovanie));

                string denotation = doc.Denotation;
                string name = doc.Naimenovanie;

                int? idTechProc = 0;
                int countTechProc = 0;
                int countNormFinish = 0;
                string norms = string.Empty;

                string remark = string.Empty;

                foreach (var techProc in doc.TechProcesses)
                {
                    if (idTechProc != techProc.ObjectID)
                    {
                        if (norms.Trim() != string.Empty)
                            norms += "\n";

                        norms += String.Format(@"Тшт = {0}, Тпз = {1} [{2} из {3}]",
                                                Math.Round(techProc.SumPieceTime.Value, 2),
                                                Math.Round(techProc.SumPrepTim.Value, 2), doc.TechProcesses.Where(t => t.ObjectID == techProc.ObjectID).Where(t => t.PieceTime != null && t.PieceTime != 0).ToList().Count,
                                                doc.TechProcesses.Where(t => t.ObjectID == techProc.ObjectID).ToList().Count);
                        idTechProc = techProc.ObjectID;
                        countTechProc++;
                    }

                    // if (techProc.PieceTime != 0 && techProc.PieceTime != null)
                    //    countNormFinish++;

                    remark = techProc.Remarks;
                }


                for (int i = 0; i < 5; i++)
                {
                    if (countTechProc == 0)
                    {
                        xls[col + i, row].Font.Color = Color.Red;
                        norms = string.Empty;
                    }
                    else
                    {
                        xls[col + i, row].Font.Color = Color.Black;
                        foreach (var tp in doc.TechProcesses)
                        {
                            if (tp.SumPieceTime == 0 || doc.TechProcesses.Where(t => t.ObjectID == tp.ObjectID).Where(t => t.PieceTime != null && t.PieceTime != 0).ToList().Count == 0)
                            {
                                xls[col + i, row].Font.Color = Color.Blue;

                            }
                        }
                    }
                }

                col = InsertCell(true, col, row, denotation.Trim(), xls);

                col = InsertCell(true, col, row, name.Trim(), xls);


                col = InsertCell(true, col, row, countTechProc.ToString(), xls);

                col = InsertCell(true, col, row, norms, xls);

                col = InsertCell(true, col, row, remark, xls);

                row++;
            }
            return row;
        }



        // Группировка элементов по типам с сортировкой внутри групп
        public static List<TFDDocument> SortData(List<TFDDocument> reportData)
        {
            List<TFDDocument> sortedReportData = new List<TFDDocument>();

            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Document));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Complex));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Assembly));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Detal));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.StandartItem));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.OtherItem));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Material));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Komplekt));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.ComplexProgram));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.ComponentProgram));

            return sortedReportData;
        }

        public static int GetSortGroupByClass(NomenclatureObject nomObject)
        {
            if (nomObject != null)
            {
                var type = nomObject.Class;

                switch (type.Guid)
                {
                    case var t when (t == Guids.Nomenclature.Types.Document):
                        return 1;
                    case var t when (t == Guids.Nomenclature.Types.Complex):
                        return 2;
                    case var t when (t == Guids.Nomenclature.Types.Assembly):
                        return 3;
                    case var t when (t == Guids.Nomenclature.Types.Detail):
                        return 4;
                    case var t when (t == Guids.Nomenclature.Types.StandartItem):
                        return 5;
                    case var t when (t == Guids.Nomenclature.Types.OtherItem):
                        return 6;
                    case var t when (t == Guids.Nomenclature.Types.Material):
                        return 7;
                    case var t when (t == Guids.Nomenclature.Types.Komplekt):
                        return 8;
                    case var t when (t == Guids.Nomenclature.Types.ComplexProgram):
                        return 9;
                    case var t when (t == Guids.Nomenclature.Types.ComponentProgram):
                        return 10;
                    default: return 11;
                }
            }
            else
                return 11;
        }

        //Сортировка внутри группы объектов одного и того же типа
        public static List<TFDDocument> SortOneTypeObjects(List<TFDDocument> reportData, int type)
        {
            List<TFDDocument> oneTypeObjects = new List<TFDDocument>();
            foreach (TFDDocument doc in reportData)
                if (doc.Class == type)
                    oneTypeObjects.Add(doc);

            oneTypeObjects.Sort(new NaimenAndOboznachComparer()); //сортируем записи по обозначению и наименованию
            return oneTypeObjects;
        }


        public static void HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            if (!reportParameters.OutputTechProcesses)
            {
                // формирование заголовков отчета сводной ведомости
                col = InsertHeader(reportParameters.Position, "Позиция", col, row, xls);
                col = InsertHeader(reportParameters.Zone, "Зона", col, row, xls);
                col = InsertHeader(reportParameters.Format, "Формат", col, row, xls);
                col = InsertHeader(reportParameters.Denotation, "Обозначение", col, row, xls);
                col = InsertHeader(reportParameters.Name, "Наименование", col, row, xls);
                col = InsertHeader(reportParameters.Count, "Количество", col, row, xls);
                col = InsertHeader(reportParameters.Remarks, "Примечание", col, row, xls);
                col = InsertHeader(reportParameters.Letter, "Литера", col, row, xls);
                col = InsertHeader(reportParameters.Id, "Id объекта", col, row, xls);
                col = InsertHeader(reportParameters.CodeRcp, "Код ОКП", col, row, xls);
                col = InsertHeader(reportParameters.FirstUse, "Первичная применяемость", col, row, xls);
                col = InsertHeader(reportParameters.CreateDate, "Дата создания", col, row, xls);
                col = InsertHeader(reportParameters.NomenclatureNumber, "Номенклатурный номер", col, row, xls);
                col = InsertHeader(reportParameters.MeasureUnit, "Единица измерения", col, row, xls);
                col = InsertHeader(reportParameters.AuthorName, "Автор", col, row, xls);
                col = InsertHeader(reportParameters.LastEditorName, "Автор последнего изменения", col, row, xls);
                if (HaveFiles != null)
                {
                    col = InsertHeader(true, "Содержит файлы", col, row, xls);
                }
            }
            else
            {
                // формирование заголовков отчета о наличии Техпроцессов

                col = InsertHeader(true, "Обозначение", col, row, xls);
                col = InsertHeader(true, "Наименование", col, row, xls);
                col = InsertHeader(true, "Кол-во техпроцессов", col, row, xls);
                col = InsertHeader(true, "Нормы [Кол-во отнормированных операций]", col, row, xls);
                col = InsertHeader(true, "Примечание", col, row, xls);

            }

            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);
        }

        public static void MakeSelectionByTypeReport(Xls xls, List<TFDDocument> reportData, ReportParameters reportParameters, LogDelegate logDelegate)
        {
            int countCol = 0;

            if (!reportParameters.OutputTechProcesses)
            {
                // формирование заголовков отчета сводной ведомости
                if (reportParameters.Position)
                    countCol++;
                if (reportParameters.Zone)
                    countCol++;
                if (reportParameters.Format)
                    countCol++;
                if (reportParameters.Denotation)
                    countCol++;
                if (reportParameters.Name)
                    countCol++;
                if (reportParameters.Count)
                    countCol++;
                if (reportParameters.Remarks)
                    countCol++;
                if (reportParameters.Letter)
                    countCol++;
                if (reportParameters.Id)
                    countCol++;
                if (reportParameters.CodeRcp)
                    countCol++;
                if (reportParameters.FirstUse)
                    countCol++;
                if (reportParameters.CreateDate)
                    countCol++;
                if (reportParameters.NomenclatureNumber)
                    countCol++;
                if (reportParameters.MeasureUnit)
                    countCol++;
                if (reportParameters.AuthorName)
                    countCol++;
                if (reportParameters.LastEditorName)
                    countCol++;

                if (HaveFiles != null)
                {
                    countCol++;
                }
            }
            else
            {
                countCol = 5;
            }

            string headerString = "Выборка по типам для устройств:";
            var listRootObj = reportData.Where(t => t.Level == 0).ToList();
            if (listRootObj.Count == 1)
            {
                headerString = "Выборка по типам для " + listRootObj[0].Naimenovanie + " " + listRootObj[0].Denotation;
            }

            countStr = 3;
            xls[1, 1].Font.Name = "Calibri";
            xls[1, 1].Font.Size = 11;
            xls[1, 1].SetValue(headerString);
            xls[1, 1].Font.Bold = true;
            xls[1, 1].Font.Underline = true;
            xls[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1, countCol, 1].Merge();

            int row = 2;
            //Формирование название отчета
            if (listRootObj.Count > 1)
            {
                InsertHeader(true, "Наименование", 1, row, xls);
                InsertHeader(true, "Обозначение", 2, row, xls);
                InsertHeader(true, "Количество", 3, row, xls);

                xls[1, 2, 3, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                           XlBordersIndex.xlEdgeBottom,
                                           XlBordersIndex.xlEdgeLeft,
                                           XlBordersIndex.xlEdgeRight,
                                           XlBordersIndex.xlInsideVertical);
                row++;
                foreach (TFDDocument doc in reportData)
                {

                    if (doc.Level == 0)
                    {
                        xls[1, row].SetValue(doc.Naimenovanie);
                        xls[2, row].SetValue(doc.Denotation);
                        foreach (var obj in reportParameters.listObjectCountId)
                        {
                            if (obj.Key == doc.ObjectID)
                            {
                                xls[3, row].SetValue(obj.Value);
                            }
                        }

                        xls[1, row, 3, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin,
                                                     XlBordersIndex.xlEdgeTop,
                                                     XlBordersIndex.xlEdgeBottom,
                                                     XlBordersIndex.xlEdgeLeft,
                                                     XlBordersIndex.xlEdgeRight,
                                                     XlBordersIndex.xlInsideVertical);
                        row++;
                    }
                }
            }

            int col = 1;
            if (!reportParameters.OutputTechProcesses || !reportParameters.GroupByDeviceAndTypes)
            {
                row++;
                HeaderTable(xls, reportParameters, 1, row);
            }



            col = countCol + 1;
            int colCount = col;
            int headerRow = 0;
            bool insertSection = false;

            row++;
            int rowMainTable = row;


            foreach (var obj in listRootObj)
            {
                if (reportData.Contains(obj) && (!reportParameters.OutputTechProcesses || (obj.Class == TFDClass.Complex && !reportParameters.ComplexTP)
                    || (obj.Class == TFDClass.Komplekt && !reportParameters.ComplementTP)
                    || (obj.Class == TFDClass.Detal && !reportParameters.DetailTP)
                    || (obj.Class == TFDClass.StandartItem && !reportParameters.StandartItemsTP)
                    || (obj.Class == TFDClass.Assembly && !reportParameters.AssemblyTP)))
                    reportData.Remove(obj);
            }
            if (reportParameters.GroupByDevice)
            {
                reportData = reportData.OrderBy(t => t.parenDevicetId).ToList();
            }
            string elementHeader = string.Empty;

            int row1;
            bool newPage = false;
            List<int> listRowForPage = new List<int>();
            List<int> listColForPage = new List<int>();
            List<int> listCountPage = new List<int>();

            if (reportParameters.OutputTechProcesses && reportParameters.GroupByDeviceAndTypes)
            {
                row++;
                var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
                var reference = nomenclatureReferenceInfo.CreateReference();
                string denotation =
                    reference.Find(reportData[0].parenDevicetId)[new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].
                        Value.ToString();
                string name = reference.Find(reportData[0].parenDevicetId)[new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")].Value.ToString();

                InsertHeader(row, col, "Проверяемый объект:", xls);
                row++;
                row1 = row;

                InsertDeviceNameHeader(row, col, name + " - " + denotation, xls);
                row++;
                row1 = row;

                HeaderTable(xls, reportParameters, 1, row);
                row++;

            }
            if (!reportParameters.OutputTechProcesses || (reportParameters.OutputTechProcesses && (reportParameters.GroupByTypes || reportParameters.GroupByDeviceAndTypes)))
            {
                if (reportData.Count > 0)
                    InsertBomSectionHeader(row, col, TFDClass.InsertSection(reportData[0].Class), xls);

                row++;
            }
            else
            {
                var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
                var reference = nomenclatureReferenceInfo.CreateReference();
                string denotation =
                    reference.Find(reportData[0].parenDevicetId)[new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].
                        Value.ToString();
                string name = reference.Find(reportData[0].parenDevicetId)[new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")].Value.ToString();

                InsertBomSectionHeader(row, col, name + " - " + denotation, xls);
                row++;
            }

            for (int i = 0; i < reportData.Count; i++)
            {

                if (reportParameters.OutputTechProcesses && reportParameters.GroupByDeviceAndTypes)
                {

                    if ((i > 0) && (reportData[i].parenDevicetId != reportData[i - 1].parenDevicetId))
                    {
                        row++;
                        InsertHeader(row, col, "Проверяемый объект:", xls);
                        row++;
                        row1 = row;


                        var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
                        var reference = nomenclatureReferenceInfo.CreateReference();
                        string denotation =
                            reference.Find(reportData[i].parenDevicetId)[
                                new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].
                                Value.ToString();

                        string name = reference.Find(reportData[i].parenDevicetId)[new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")].Value.ToString();
                        InsertDeviceNameHeader(row, col, name + " - " + denotation, xls);
                        countStr++;
                        row++;

                        row1 = row;

                        HeaderTable(xls, reportParameters, 1, row);
                        row++;
                        insertSection = true;

                    }
                    else
                    {
                        insertSection = false;
                    }


                }

                row1 = row;
                xls[1, row1, colCount - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                       XlBordersIndex.xlEdgeBottom,
                                                                                                                       XlBordersIndex.xlEdgeLeft,
                                                                                                                       XlBordersIndex.xlEdgeRight,
                                                                                                                       XlBordersIndex.xlInsideVertical);
                if (!reportParameters.OutputTechProcesses || (reportParameters.OutputTechProcesses && (reportParameters.GroupByTypes || reportParameters.GroupByDeviceAndTypes)))
                {
                    if (((i > 0) && (reportData[i].Class != reportData[i - 1].Class)) || insertSection)
                    {

                        InsertBomSectionHeader(row, col, TFDClass.InsertSection(reportData[i].Class), xls);
                        countStr++;
                        row++;
                        row1 = row;
                        xls[1, row1, colCount - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                               XlBordersIndex.xlEdgeBottom,
                                                                                                                               XlBordersIndex.xlEdgeLeft,
                                                                                                                               XlBordersIndex.xlEdgeRight,
                                                                                                                               XlBordersIndex.xlInsideVertical);
                    }
                }
                else
                {
                    if ((i > 0) && (reportData[i].parenDevicetId != reportData[i - 1].parenDevicetId))
                    {
                        var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
                        var reference = nomenclatureReferenceInfo.CreateReference();
                        string denotation =
                            reference.Find(reportData[i].parenDevicetId)[
                                new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].
                                Value.ToString();

                        string name = reference.Find(reportData[i].parenDevicetId)[new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")].Value.ToString();

                        InsertBomSectionHeader(row, col, name + " - " + denotation, xls);
                        countStr++;
                        row++;
                    }
                }


                if (!reportParameters.OutputTechProcesses)
                {
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Documentation,
                                       TFDClass.Document, xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Complex, TFDClass.Complex,
                                       xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Assembly,
                                       TFDClass.Assembly, xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Detail, TFDClass.Detal,
                                       xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Standart,
                                       TFDClass.StandartItem, xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Other, TFDClass.OtherItem,
                                       xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Materials,
                                       TFDClass.Material, xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.Complement,
                                       TFDClass.Komplekt, xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.ComplexProgram,
                                       TFDClass.ComplexProgram, xls, logDelegate);
                    row = InsertObject(reportParameters, reportData[i], row, reportParameters.ComponentProgram,
                                       TFDClass.ComponentProgram, xls, logDelegate);
                }
                else
                {

                    // else
                    {
                        row = InsertObjectTP(reportParameters, reportData[i], row, reportParameters.ComplexTP,
                                       TFDClass.Complex,
                                       xls, logDelegate);
                        row = InsertObjectTP(reportParameters, reportData[i], row, reportParameters.AssemblyTP,
                                           TFDClass.Assembly, xls, logDelegate);
                        row = InsertObjectTP(reportParameters, reportData[i], row, reportParameters.DetailTP, TFDClass.Detal,
                                           xls, logDelegate);
                        row = InsertObjectTP(reportParameters, reportData[i], row, reportParameters.StandartItemsTP,
                                           TFDClass.StandartItem, xls, logDelegate);
                        row = InsertObjectTP(reportParameters, reportData[i], row, reportParameters.ComplementTP,
                                           TFDClass.Komplekt, xls, logDelegate);

                    }

                }

                if (reportParameters.Main5Col)
                    if (countStr > 33)
                    {

                        countPage++;
                        int countRowDec = 2;

                        // Для случая когда  первая строка страницы - заголовок раздела
                        if (((countStr == 35) && (countWrap == 1)) || ((countStr == 36) && (countWrap == 2)) || ((countStr == 37) && (countWrap == 3))
                            || ((countStr == 38) && (countWrap == 4)) || ((countStr == 39) && (countWrap == 5)))
                        {
                            countRowDec = countWrap + 2;
                        }



                        countStr = countWrap;

                        listColForPage.Add(col);
                        listRowForPage.Add(row - countRowDec);
                        listCountPage.Add(countPage);
                        //col--;

                        newPage = true;
                    }
                    else
                    {
                        newPage = false;
                    }
            }
            if (!reportParameters.GroupByDeviceAndTypes)
            {

                xls[1, rowMainTable, colCount - 1, row - rowMainTable].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                                             XlBordersIndex.xlEdgeRight,
                                                                                                             XlBordersIndex.xlInsideVertical);
            }
            xls[1, 1, colCount - 1, row - 1].VerticalAlignment = XlVAlign.xlVAlignTop;

            if (reportParameters.Main5Col)
            {
                for (int i = 0; i < listColForPage.Count; i++)
                {
                    if (listColForPage.Count > 1)

                        InsertCell(listColForPage[i], listRowForPage[i],
                            "с" + listCountPage[i].ToString() + "из" +
                            (listCountPage[listCountPage.Count - 1] + 1).ToString(), xls);
                }

                if (listColForPage.Count > 1)
                    InsertCell(col, row - 1,
                        "с" + (listCountPage[listCountPage.Count - 1] + 1) + "из" +
                        (listCountPage[listCountPage.Count - 1] + 1), xls);
            }

        }

        // Сложение количеств повторяющихся элементов отсортированной коллекции
        static private List<TFDDocument> SumDublicates(List<TFDDocument> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].ObjectID == data[slaveID].ObjectID && data[mainID].Level != 0 && data[slaveID].Level != 0)
                    {
                        data[mainID].Amount += data[slaveID].Amount;
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
            return data;
        }

        static private List<TFDDocument> SumDublicatesWithMeasureUnit(List<TFDDocument> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].ObjectID == data[slaveID].ObjectID &&
                        string.Join("; ", data[mainID].MeasureUnit.Where(t => t.Trim() != string.Empty)) ==
                        string.Join("; ", data[slaveID].MeasureUnit.Where(t => t.Trim() != string.Empty)) &&
                        data[mainID].Level != 0 && data[slaveID].Level != 0)
                    {
                        data[mainID].Amount += data[slaveID].Amount;
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
            return data;
        }
    }

    internal class NaimenAndOboznachComparer : IComparer<TFDDocument>
    {
        public int Compare(TFDDocument ob1, TFDDocument ob2)
        {
            string designation1 = ob1.Denotation.Replace(" ", "");
            string designation2 = ob2.Denotation.Replace(" ", "");
            int ob = String.Compare(designation1, designation2);
            if (ob != 0)
                return ob;
            else
                return String.Compare(ob1.Naimenovanie, ob2.Naimenovanie);
        }
    }


}
public class TextFormatter
{
    Font _font;

    int _displayResolutionX;

    public TextFormatter(Font font)
    {
        _font = font;
        _displayResolutionX = GetDisplayResolutionX();
    }

    [DllImport("User32.dll")]
    static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("gdi32.dll")]
    static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    public enum DeviceCap
    {
        #region
        /// <summary>
        /// Device driver version
        /// </summary>
        DRIVERVERSION = 0,
        /// <summary>
        /// Device classification
        /// </summary>
        TECHNOLOGY = 2,
        /// <summary>
        /// Horizontal size in millimeters
        /// </summary>
        HORZSIZE = 4,
        /// <summary>
        /// Vertical size in millimeters
        /// </summary>
        VERTSIZE = 6,
        /// <summary>
        /// Horizontal width in pixels
        /// </summary>
        HORZRES = 8,
        /// <summary>
        /// Vertical height in pixels
        /// </summary>
        VERTRES = 10,
        /// <summary>
        /// Number of bits per pixel
        /// </summary>
        BITSPIXEL = 12,
        /// <summary>
        /// Number of planes
        /// </summary>
        PLANES = 14,
        /// <summary>
        /// Number of brushes the device has
        /// </summary>
        NUMBRUSHES = 16,
        /// <summary>
        /// Number of pens the device has
        /// </summary>
        NUMPENS = 18,
        /// <summary>
        /// Number of markers the device has
        /// </summary>
        NUMMARKERS = 20,
        /// <summary>
        /// Number of fonts the device has
        /// </summary>
        NUMFONTS = 22,
        /// <summary>
        /// Number of colors the device supports
        /// </summary>
        NUMCOLORS = 24,
        /// <summary>
        /// Size required for device descriptor
        /// </summary>
        PDEVICESIZE = 26,
        /// <summary>
        /// Curve capabilities
        /// </summary>
        CURVECAPS = 28,
        /// <summary>
        /// Line capabilities
        /// </summary>
        LINECAPS = 30,
        /// <summary>
        /// Polygonal capabilities
        /// </summary>
        POLYGONALCAPS = 32,
        /// <summary>
        /// Text capabilities
        /// </summary>
        TEXTCAPS = 34,
        /// <summary>
        /// Clipping capabilities
        /// </summary>
        CLIPCAPS = 36,
        /// <summary>
        /// Bitblt capabilities
        /// </summary>
        RASTERCAPS = 38,
        /// <summary>
        /// Length of the X leg
        /// </summary>
        ASPECTX = 40,
        /// <summary>
        /// Length of the Y leg
        /// </summary>
        ASPECTY = 42,
        /// <summary>
        /// Length of the hypotenuse
        /// </summary>
        ASPECTXY = 44,
        /// <summary>
        /// Shading and Blending caps
        /// </summary>
        SHADEBLENDCAPS = 45,

        /// <summary>
        /// Logical pixels inch in X
        /// </summary>
        LOGPIXELSX = 88,
        /// <summary>
        /// Logical pixels inch in Y
        /// </summary>
        LOGPIXELSY = 90,

        /// <summary>
        /// Number of entries in physical palette
        /// </summary>
        SIZEPALETTE = 104,
        /// <summary>
        /// Number of reserved entries in palette
        /// </summary>
        NUMRESERVED = 106,
        /// <summary>
        /// Actual color resolution
        /// </summary>
        COLORRES = 108,

        // Printing related DeviceCaps. These replace the appropriate Escapes
        /// <summary>
        /// Physical Width in device units
        /// </summary>
        PHYSICALWIDTH = 110,
        /// <summary>
        /// Physical Height in device units
        /// </summary>
        PHYSICALHEIGHT = 111,
        /// <summary>
        /// Physical Printable Area x margin
        /// </summary>
        PHYSICALOFFSETX = 112,
        /// <summary>
        /// Physical Printable Area y margin
        /// </summary>
        PHYSICALOFFSETY = 113,
        /// <summary>
        /// Scaling factor x
        /// </summary>
        SCALINGFACTORX = 114,
        /// <summary>
        /// Scaling factor y
        /// </summary>
        SCALINGFACTORY = 115,

        /// <summary>
        /// Current vertical refresh rate of the display device (for displays only) in Hz
        /// </summary>
        VREFRESH = 116,
        /// <summary>
        /// Horizontal width of entire desktop in pixels
        /// </summary>
        DESKTOPVERTRES = 117,
        /// <summary>
        /// Vertical height of entire desktop in pixels
        /// </summary>
        DESKTOPHORZRES = 118,
        /// <summary>
        /// Preferred blt alignment
        /// </summary>
        BLTALIGNMENT = 119
        #endregion
    }

    int GetDisplayResolutionX()
    {
        IntPtr p = GetDC(new IntPtr(0));
        int res = GetDeviceCaps(p, (int)DeviceCap.LOGPIXELSX);
        return res;
    }

    /// Коэффициент, учитывающий различие между шириной строки,
    /// определяемой в этой программе и шириной строки, получаемой в CAD
    static readonly float widthFactor = 1.5f;

    public double GetTextWidth(string text, Font font)
    {
        Size size = TextRenderer.MeasureText(text, font);
        double width = 25.4f * widthFactor * (double)size.Width / (double)_displayResolutionX;
        return width;
    }

    public string[] Wrap(string text, double maxWidth)
    {
        return Wrap(text, _font, maxWidth);
    }

    /// Разбивка на строки по символам перевода строки \n
    string[] Wrap(string text, Font font, double maxWidth)
    {
        string[] lines = text.Split(new string[] { @"\n" }, StringSplitOptions.None);


        List<string> wrappedLines = new List<string>();
        foreach (string line in lines)
        {
            string[] wrappedBySpaces = WrapBySyllables(line, font, maxWidth);
            wrappedLines.AddRange(wrappedBySpaces);
        }
        return wrappedLines.ToArray();
    }

    /// Выделение слога - последовательности символов с завершающим пробелом
    /// или символом мягкого переноса. Символ мягкого переноса (если присутствует)
    /// выносится в отдельную группу
    static readonly Regex _syllableRegex = new Regex(
        @"(?<syllable>((ГОСТ|ОСТ)[\x20\u00AD]+)?[^\x20\u00AD]+)(?<softHyphen>\u00AD+)?",
        RegexOptions.Compiled);

    /// Разбивка на строки по слогам (разделители - пробелы и знаки мягкого переноса)
    string[] WrapBySyllables(string text, Font font, double maxWidth)
    {
        if (text == "")
            return new string[] { "" };

        List<string> lines = new List<string>();
        MatchCollection mc = _syllableRegex.Matches(text);
        string currentLine = "";
        bool currentLineEndsWithSoftHyphen = false;
        foreach (Match match in mc)
        {
            string syllable = match.Groups["syllable"].Value;
            bool endsBySoftHyphen = match.Groups["softHyphen"].Success;

            string candidateLine;
            if (currentLine.Length > 0)
            {
                if (currentLineEndsWithSoftHyphen)
                    candidateLine = currentLine + syllable;
                else
                    candidateLine = currentLine + " " + syllable;
            }
            else
                candidateLine = syllable;
            double candidateLineWidth = GetTextWidth(candidateLine, font);

            if (candidateLineWidth > maxWidth)
            {
                // Перенос очередной последовательности символов на новую строку

                if (currentLine.Length > 0)
                {
                    if (currentLineEndsWithSoftHyphen)
                        lines.Add(currentLine + "-");
                    else
                        lines.Add(currentLine);
                }
                currentLine = "";
                currentLineEndsWithSoftHyphen = false;

                candidateLine = syllable;
                candidateLineWidth = GetTextWidth(candidateLine, font);
            }

            if (candidateLineWidth > maxWidth)
            {
                // Разбивка строки-кандидата по разделителям-не буквам

                if (currentLine.Length > 0)
                {
                    if (currentLineEndsWithSoftHyphen)
                        lines.Add(currentLine + "-");
                    else
                        lines.Add(currentLine);
                }
                currentLine = "";
                currentLineEndsWithSoftHyphen = false;

                string[] candidateLines = WrapByAlphaAndDelimiter(candidateLine, font, maxWidth);
                foreach (string s in candidateLines)
                    lines.Add(s);
            }
            else
            {
                // строка-кандидат не превышает максимально допустимую ширину
                currentLine = candidateLine;
                currentLineEndsWithSoftHyphen = endsBySoftHyphen;
            }
        }
        if (currentLine.Length > 0)
            lines.Add(currentLine);
        return lines.ToArray();
    }

    static readonly Regex _alphaAndDelimiterRegex = new Regex(
        @"[\d()*\p{Ll}\p{Lu}]+[^\p{Ll}\p{Lu}]*", RegexOptions.Compiled);

    /// Разбивка на строки по разделителям-не буквам
    string[] WrapByAlphaAndDelimiter(string text, Font font, double maxWidth)
    {
        if (text == "")
            return new string[] { "" };

        List<string> lines = new List<string>();
        MatchCollection mc = _alphaAndDelimiterRegex.Matches(text);
        string currentLine = "";
        foreach (Match match in mc)
        {
            string candidateLine = currentLine + match.Value;
            double candidateLineWidth = GetTextWidth(candidateLine, font);

            if (candidateLineWidth > maxWidth)
            {
                // Перенос очередной последовательности символов на новую строку

                if (currentLine.Length > 0)
                    lines.Add(currentLine);
                currentLine = "";

                candidateLine = match.Value;
                candidateLineWidth = GetTextWidth(candidateLine, font);
            }

            if (candidateLineWidth > maxWidth)
            {
                // Жесткая разбивка новой последовательности символов

                if (currentLine.Length > 0)
                    lines.Add(currentLine);
                currentLine = "";

                string[] candidateLines = WrapByCharacters(candidateLine, font, maxWidth);
                foreach (string s in candidateLines)
                    lines.Add(s);
            }
            else
                // строка-кандидат не превышает максимально допустимую ширину
                currentLine = candidateLine;
        }
        if (currentLine.Length > 0)
            lines.Add(currentLine);
        return lines.ToArray();
    }

    /// Разбивка на строки по символам
    string[] WrapByCharacters(string text, Font font, double maxWidth)
    {
        if (text == "")
            return new string[] { "" };

        List<string> lines = new List<string>();
        string currentLine = "";
        for (int i = 0; i < text.Length; ++i)
        {
            string candidateLine = currentLine + text.Substring(i, 1);
            double candidateLineWidth = GetTextWidth(candidateLine, font);
            if (candidateLineWidth > maxWidth)
            {
                if (currentLine.Length > 0)
                    lines.Add(currentLine);
                currentLine = text.Substring(i, 1);
            }
            else
                currentLine = candidateLine;
        }
        if (currentLine.Length > 0)
            lines.Add(currentLine);
        return lines.ToArray();
    }
}

public static class TFDClass
{
    public static readonly int Assembly = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Сборочная единица").Id;
    public static readonly int Detal = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Деталь").Id;
    public static readonly int Document = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find(new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114")).Id;
    public static readonly int Komplekt = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Комплект").Id;
    public static readonly int Complex = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Комплекс").Id;
    public static readonly int OtherItem = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Прочее изделие").Id;
    public static readonly int StandartItem = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Стандартное изделие").Id;
    public static readonly int Material = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find(new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987")).Id;
    public static readonly int ComponentProgram = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Компонент (программы)").Id;
    public static readonly int ComplexProgram = new NomenclatureReference(ServerGateway.Connection).Classes.AllClasses.Find("Комплекс (программы)").Id;

    public static readonly string ComplementName = "Комплекты";
    public static readonly string OtherItemsName = "Прочие изделия";
    public static readonly string StandartItemsName = "Стандартные изделия";
    public static readonly string AssemblyName = "Сборочные единицы";
    public static readonly string DetailName = "Детали";
    public static readonly string MaterialName = "Материалы";
    public static readonly string DocumentName = "Документы";
    public static readonly string ComplexName = "Комплексы";
    public static readonly string ComponentProgramName = "Компоненты (ЕСПД)";
    public static readonly string ComplexProgramName = "Комплексы (ЕСПД)";

    public static string InsertSection(int classId)
    {
        string sectionName = string.Empty;
        if (classId == Komplekt)
            sectionName = ComplementName;
        if (classId == OtherItem)
            sectionName = OtherItemsName;
        if (classId == StandartItem)
            sectionName = StandartItemsName;
        if (classId == Assembly)
            sectionName = AssemblyName;
        if (classId == Detal)
            sectionName = DetailName;
        if (classId == Material)
            sectionName = MaterialName;
        if (classId == Document)
            sectionName = DocumentName;
        if (classId == Complex)
            sectionName = ComplexName;
        if (classId == ComponentProgram)
            sectionName = ComponentProgramName;
        if (classId == ComplexProgram)
            sectionName = ComplexProgramName;

        return sectionName;
    }
}

