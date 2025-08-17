using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Procedures;
using TFlex.DOCs.Model.References.Processes;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using Filter = TFlex.DOCs.Model.Search.Filter;

namespace ExecutorProcessReport
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

        SelectionForm selectionForm = new SelectionForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            try
            {
                var reportObject = GetReportObject(context);
                var proceduresByObject = GetProceduresByObject(reportObject);

                //if (reportObject.Class.IsInherit(Guids.TypeOrderRkk))
                //{
                //    var reportParameters = new ReportParameters();
                //    reportParameters.IsAllObjects = true;
                //    reportParameters.IsTypeOrder = true;
                //    reportParameters.ProcRefObj = proceduresByObject.First();
                    
                //    var reportClass = new ExceedingTermsReport(reportParameters);
                //}
                //else
                {
                    using (new WaitCursorHelper(false))
                    {
                        selectionForm.Init(reportObject, proceduresByObject);
                        if (selectionForm.ShowDialog() == DialogResult.OK)
                        {
                            var reportClass = new ExceedingTermsReport(selectionForm.reportParameters);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        private ReferenceObject GetReportObject(IReportGenerationContext context)
        {
            //Получаем справочник выделенного объекта
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(context.ObjectsInfo[0].ReferenceId);
            var reference = referenceInfo.CreateReference();

            //Находим выделенный объект
            var filter = new Filter(referenceInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, context.ObjectsInfo[0].ObjectId);
            return reference.Find(filter).FirstOrDefault();
        }
        private List<ProcedureReferenceObject> GetProceduresByObject(ReferenceObject reportObject)
        {
            var proceduresByType = new List<ProcedureReferenceObject>();
            var referenceProcedureInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Процедуры.id);
            var referenceProcedure = referenceProcedureInfo.CreateReference();

            var filterProcedure = new Filter(referenceProcedureInfo);
            filterProcedure.Terms.AddTerm("[" + Guids.Процедуры.СписокОбъектов.Справочники.Id + "]->[" + Guids.Процедуры.СписокОбъектов.Справочники.Поля.Справочник + "]", ComparisonOperator.Equal, reportObject.Reference.ParameterGroup.Guid);

            var proceduresBySelectedReference = referenceProcedure.Find(filterProcedure).ToList()
               .Select(t => (ProcedureReferenceObject)t)
               .Where(t => Convert.ToInt32(t[Guids.Процедуры.Поля.ДоступНаЗапуск].Value) == 0 || (Convert.ToInt32(t[Guids.Процедуры.Поля.ДоступНаЗапуск].Value) == 1 && Convert.ToInt32(t.GetObjects(Guids.Процедуры.Связи.ПравоНаЗапускБп).Count) > 0))
               .Where(t => !t.Name.ToString().Trim().ToLower().Contains("аннулирование")
               && !t.Name.ToString().ToLower().Trim().Contains("доступ")
               && !t.Name.ToString().ToLower().Trim().Contains("макрос"))
               .Distinct()
               .ToList();

            foreach (var procedure in proceduresBySelectedReference.Distinct())
            {
                var conditionsProcedure = procedure.GetObjects(Guids.Процедуры.СписокОбъектов.Справочники.Id).Select(r => r[Guids.Процедуры.СписокОбъектов.Справочники.Поля.Условие].Value.ToString());
                var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(reportObject.Reference.Id);
                var reference = referenceInfo.CreateReference();
                foreach (var condition in conditionsProcedure)
                {
                    if (condition.Contains(reportObject.Class.ToString()))
                    {
                        proceduresByType.Add(procedure);
                        //break;
                    }
                }
            }
            return proceduresByType;
        }
    }

    public delegate void LogDelegate(string line);
    public class ExceedingTermsReport
    {
        public IndicationForm m_form = new IndicationForm();
        public void Dispose()
        {
        }

        public ExceedingTermsReport(ReportParameters reportParameters)
        {
            Make(reportParameters);
        }

        public void Make(ReportParameters reportParameters)
        {
            var m_writeToLog = new LogDelegate(m_form.writeToLog);
            m_form.TopMost = true;
            m_form.Visible = true;
            m_form.Activate();
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            var reportData = GetObjectsBySettings(reportParameters, m_writeToLog);

            if (reportData.Count == 0)
            {
                this.Dispose();
                m_form.Close();
                return;
            }

            var xls = new Xls();

            try
            {          
                xls.Application.DisplayAlerts = false;
                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");
                //Формирование отчета
                MakeSelectionByTypeReport(xls, reportParameters, reportData, m_writeToLog, m_form.progressBar);
                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                xls.Application.DisplayAlerts = true;
                xls.Visible = true;
            }
            m_form.Dispose();
        }

        private List<ActionParameters> GetObjectsBySettings(ReportParameters reportParameters, LogDelegate m_writeToLog)
        {
            m_form.setStage(IndicationForm.Stages.DataAcquisition);
            var reportData = new List<ActionParameters>();
            m_writeToLog("=== Обработка данных ===");

            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Процессы.Id);
            var reference = referenceInfo.CreateReference();
            var filter = new Filter(referenceInfo);
            filter.Terms.AddTerm("[" + Guids.Процессы.Связи.Процедура + "]->[" + Guids.Процедуры.Поля.Наименование + "]", ComparisonOperator.Equal, reportParameters.ProcRefObj.Name);

            if (reportParameters.IsCurrentState)
            {
                filter.Terms.AddTerm("[" + Guids.Процессы.Поля.Статус + "]", ComparisonOperator.NotEqual, 4);
            }

            if (reportParameters.IsSelectedState || reportParameters.IsAllState)
            {
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.CreationDate], ComparisonOperator.GreaterThan, reportParameters.BeginPeriod);
                if (reportParameters.IsSelectedState)
                {
                    filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.CreationDate], ComparisonOperator.LessThan, reportParameters.EndPeriod);
                }
            }

            var findedObjects = reference.Find(filter).ToList();
            var findedObjectsBySettings = GetObjectsByDepartmentOrAuthor(reportParameters, findedObjects);

            if (findedObjectsBySettings.Count == 0)
            {
                MessageBox.Show("Количество объектов для отчета 0. ", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            var progressStep = Convert.ToDouble(m_form.progressBar.Maximum) / (findedObjectsBySettings.Count * 2);

            if (m_form.progressBar.Maximum < (findedObjectsBySettings.Count * 2))
            {
                m_form.progressBar.Step = 1;
            }
            else
            {
                m_form.progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            foreach (var obj in findedObjectsBySettings.Select(t => (WorkflowReferenceObject)t))
            {
                var rkk = obj.GetObjects(Guids.Процессы.Связи.ЗапущенныеОбъекты).ToList().FirstOrDefault();
                if (rkk == null)
                {
                    continue;
                }

                if (reportParameters.IsSelectedState || reportParameters.IsAllState)
                {
                    var параметрОт = (DateTime)rkk[Guids.Ркк.Поля.От];
                    if (параметрОт < reportParameters.BeginPeriod || параметрОт > reportParameters.EndPeriod)
                    {
                        continue;
                    }
                }

                m_writeToLog(string.Format("Обработка объекта: " + (rkk != null ? rkk.ToString() : string.Empty) + " " + obj.SystemFields.CreationDate + " " + obj.SystemFields.Author));
                m_form.progressBar.PerformStep();

                var makerComments = ReadData(obj, reportParameters);

                if (makerComments != null)
                {
                    foreach (var mC in makerComments)
                    {
                        if (mC != null)
                        {
                            foreach (var act in mC.Actions)
                            {
                                act.RKK = rkk;
                                reportData.Add(act);
                            }
                        }
                    }
                }
            }
            return reportParameters.IsControlState ? reportData.Where(t => СрокСогласованияПревышен(t.StartTime, t.EndTime, reportParameters.CountControlDays)).ToList() :
                reportParameters.Chiefs ? reportData.Where(t => руководители.Contains(t.MakerName)).ToList() : reportData;
        }

        private List<string> руководители = new List<string>()
        {
            "Бугук К.С.",
            "Гоева О.Н.",
            "Заводчиков Ю.Л.",
            "Кириллов В.Н.",
            "Михалев А.О.",
            "Петрунин В.В.",
            "Ступнев В.Ю.",
            "Трубников А.А.",
            "Федоров А.М."
        };

        private List<ReferenceObject> GetObjectsByDepartmentOrAuthor(ReportParameters reportParameters, List<ReferenceObject> findedObjects)
        {
            var findedObjectsBySettings = new List<ReferenceObject>();

            if (reportParameters.IsCurrentState || reportParameters.IsAllState)
            {
                if (reportParameters.IsObjectsByDepartment && reportParameters.Department != null)
                {
                    foreach (var process in findedObjects)
                    {
                        try
                        {
                            if (process.GetObjects(Guids.Процессы.Связи.ЗапущенныеОбъекты).Any(t => t.Class.IsInherit(Guids.Ркк.Типы.РегистрационноКонтрольныеКарточки) &&
                            t.GetObject(Guids.Ркк.Связи.Подразделение).SystemFields.Id == reportParameters.Department.SystemFields.Id))
                            {
                                findedObjectsBySettings.Add(process);
                            }
                        }
                        catch
                        {

                        }
                    }
                }
                else
                {
                    if (reportParameters.IsObjectsByAuthor && reportParameters.Author != null)
                    {
                        foreach (var process in findedObjects)
                        {
                            if (process.GetObjects(Guids.Процессы.Связи.ЗапущенныеОбъекты).Any(t => t.Class.IsInherit(Guids.Ркк.Типы.РегистрационноКонтрольныеКарточки) &&
                            t.SystemFields.AuthorId == reportParameters.Author.SystemFields.Id))
                            {
                                findedObjectsBySettings.Add(process);
                            }
                        }
                    }
                    else
                    {
                        if (reportParameters.IsAllObjects)
                        {
                            return findedObjects;
                        }
                    }
                }
            }
            else
            {
                return findedObjects;
            }
            return findedObjectsBySettings;
        }

        public static int InsertHeader(string header, int col, int row, Xls xls)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row].Interior.Color = Color.LightGray;
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].WrapText = true;

            col++;

            return col;
        }

        public int PasteHeader(Xls xls, ReportParameters reportParameters, int col, int row)
        {

            col = InsertHeader("№ РКК", col, row, xls);
            if (reportParameters.IsCurrentState || reportParameters.IsAllState)
            {
                col = InsertHeader("Этап", col, row, xls);
            }
            col = InsertHeader("Исполнитель", col, row, xls);
            col = InsertHeader("Время на этапе", col, row, xls);

            if (reportParameters.IsTypeOrder && reportParameters.IsCurrentState)
            {
                col = InsertHeader("Заказчик", col, row, xls);
                col = InsertHeader("Шифр заказа", col, row, xls);
            }
            else
            {
               
                if (reportParameters.Chiefs)
                {
                    col = InsertHeader("Содержание", col, row, xls);
                    col = InsertHeader("Контрагент", col, row, xls);
                    col = InsertHeader("Сумма", col, row, xls);
                }
                else
                {
                    col = !reportParameters.IsControlState ? InsertHeader("Фактическое время на этапе", col, row, xls) :
                    InsertHeader("Количество дней на этапе", col, row, xls);
                    col = InsertHeader("Комментарий", col, row, xls);
                }
            }
           
            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,                                                                                      XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            row++;
            return row;
        }

        public List<MakerComment> ReadData(WorkflowReferenceObject processObject,  ReportParameters reportParameters)
        {
            var errorId = 0;
            List<MakerComment> stateActionsObjects = new List<MakerComment>();
            try
            {
                errorId = 5;
                //Все состояния
                var states = reportParameters.IsSelectedState ?
                    processObject.Procedure.GetObjects(Guids.Процедуры.СписокОбъектов.Состояния.Id).Where(t => t.ToString() == reportParameters.StatesObject.First().ToString()).ToList() :
                    processObject.Procedure.GetObjects(Guids.Процедуры.СписокОбъектов.Состояния.Id).ToList();

                errorId = 7;
                if (reportParameters.IsSelectedState)
                {
                    var actions = processObject.Links.ToMany[Guids.Процессы.Связи.Действия].Objects
                    .Where(t => Convert.ToDateTime(t[Guids.Процедуры.СписокОбъектов.Состояния.Поля.ВремяЗавершения].Value) > reportParameters.BeginPeriod && Convert.ToDateTime(t[Guids.Процедуры.СписокОбъектов.Состояния.Поля.ВремяЗавершения].Value) < reportParameters.EndPeriod)
                    .ToList();
                    errorId = 4;
                    //Последнее действие процесса
                  //  var lastAction = actions.Last();
                    stateActionsObjects.AddRange(GetAllMakerComment(actions, states));
                }

                if (reportParameters.IsCurrentState)
                {
                    var activeActions = processObject.GetObjects(Guids.Процессы.Связи.ТекукщиеДействия);
                    var actions2 = processObject.GetObjects(Guids.Процессы.Связи.Действия).ToList();
                    var currentActions = actions2.Where(t => t.Class.IsInherit(Guids.ДействияПроцессов.Типы.Состояния))
                        .Where(t => activeActions.Select(a => (Guid)a[Guids.ТекущиеДействияПроцессов.Поля.ИдентификаторЭтапаПроцесса].Value)
                        .Contains((Guid)t[Guids.ДействияПроцессов.Поля.ИдентификаторСостояния].Value))
                        .Where(t => string.IsNullOrEmpty(t[Guids.Процедуры.СписокОбъектов.Состояния.Поля.ВремяЗавершения].ToString())).ToList();
                    if (currentActions.Count > 0)
                    {
                        stateActionsObjects.AddRange(GetAllMakerComment(currentActions, states));
                    }
                }
              
                if (reportParameters.IsAllState)
                {
                    var startedObject = processObject.GetObjects(Guids.Процессы.Связи.ЗапущенныеОбъекты).First();
                    if (startedObject.Class.IsInherit(Guids.Ркк.Типы.РегистрационноКонтрольныеКарточки) &&
                        ((DateTime)startedObject[Guids.Ркк.Поля.От]) > reportParameters.BeginPeriod &&
                        ((DateTime)startedObject[Guids.Ркк.Поля.От]) < reportParameters.EndPeriod)
                    {
                        var actions = processObject.GetObjects(Guids.Процессы.Связи.Действия).ToList();
                        stateActionsObjects.AddRange(GetAllMakerComment(actions, states));
                    }
                   
                }

                errorId = 8;
            }
            catch
            {
                var processStartObjects = processObject.GetObjects(Guids.Процессы.Связи.ЗапущенныеОбъекты).ToList();
                MessageBox.Show("Ошибка № " + errorId + "\r\n в объекте " + string.Join("\r\n", processStartObjects.Select(t=>t.ToString())) , "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return stateActionsObjects;
        }

        public static List<MakerComment> GetAllMakerComment(List<ReferenceObject> actions, List<ReferenceObject> states)
        {
            List<MakerComment> listMakerComment = new List<MakerComment>();
            //Действия работа или согласование

            var workActions = actions.Where(t => t.Class.IsInherit(Guids.ДействияПроцессов.Типы.Состояния) && states.Where(s => s.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Работа)
                                                               || s.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Согласование)).
                                                                           Select(s => s.SystemFields.Guid).
                                            Contains((Guid)t[Guids.ДействияПроцессов.Поля.ИдентификаторСостояния].Value)).ToList();
            if (workActions.Count() > 0)
            {
                var statesWork = states.Where(t => workActions.Select(a => ((Guid)a[Guids.ДействияПроцессов.Поля.ИдентификаторСостояния].Value)).
                                                   Contains((t.SystemFields.Guid))).
                                                   Where(t => t.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Работа) || t.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Согласование)).ToList();
                //Добавление в список найденных пользователей и их комментариев
                listMakerComment.AddRange(GetListMakerComment(actions, statesWork, workActions, 1));
            }
            return listMakerComment;
        }

        //Добавление в список найденных пользователей и их комментариев
        private static List<MakerComment> GetListMakerComment(List<ReferenceObject> actions, List<ReferenceObject> statesByType, List<ReferenceObject> searchedActions, byte workType)
        {
            Dictionary<ReferenceObject, List<ReferenceObject>> stateActionsByType = new Dictionary<ReferenceObject, List<ReferenceObject>>();
            foreach (var state in statesByType)
            {
                stateActionsByType.Add(state, searchedActions.Where(t => t.Class.IsInherit(Guids.ДействияПроцессов.Типы.Состояния)).Where(t => ((Guid)t[Guids.ДействияПроцессов.Поля.ИдентификаторСостояния].Value) == state.SystemFields.Guid).ToList());
            }
            return stateActionsByType.Select(t => new MakerComment(t, t.Value.Select(d => d.SystemFields.CreationDate).ToList(), workType)).ToList();

        }

        //Рекурсивный поиск действий процессов
        public static List<ReferenceObject> GetActionToWorkState(List<ReferenceObject> currentActions, List<ReferenceObject> statesByType)
        {
            List<ReferenceObject> listActions = new List<ReferenceObject>();
            foreach (var action in currentActions)
            {
                listActions.AddRange(action.Parents);
                listActions.AddRange(GetActionToWorkState(action.Parents.Where(t => !statesByType.Select(s => s.SystemFields.Guid).Contains((Guid)t[Guids.ДействияПроцессов.Поля.ИдентификаторСостояния].Value)).ToList(), statesByType));
            }
            return listActions;
        }

        public bool СрокСогласованияПревышен(DateTime startTime, DateTime endTime, int нормативныйСрок)
        {
            var user = ClientView.Current.GetUser();
            var calendar = user.Calendar;
            return endTime != DateTime.MinValue ? calendar.WorkTimeManager.CalcDurationInDays(startTime, endTime) >= нормативныйСрок + 1 :
                calendar.WorkTimeManager.CalcDurationInDays(startTime, DateTime.Now.Date) >= нормативныйСрок + 1;
        }
        public int ПолучитьЗадержкуВДнях (DateTime startTime, DateTime endTime, int нормативныйСрок)
        {
            var user = ClientView.Current.GetUser();
            var calendar = user.Calendar;
            return endTime != DateTime.MinValue ? calendar.WorkTimeManager.CalcDurationInDays(startTime, endTime) - (нормативныйСрок) :
                calendar.WorkTimeManager.CalcDurationInDays(startTime, DateTime.Now.Date) - (нормативныйСрок);
        }
        public TimeSpan ПолучитьЗадержкуВЧасах(DateTime startTime, DateTime endTime)
        {
            var user = ClientView.Current.GetUser();
            var calendar = user.Calendar;
            return endTime != DateTime.MinValue ? calendar.WorkTimeManager.CalcTaskDuration(startTime, endTime) :
                calendar.WorkTimeManager.CalcTaskDuration(startTime, DateTime.Now.Date);
        }
        private void MakeSelectionByTypeReport(Xls xls, ReportParameters reportParameters, List<ActionParameters> reportData, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            var progressStep = Convert.ToDouble(progressBar.Maximum) / (reportData.Count * 2);

            if (progressBar.Maximum < (reportData.Count * 2))
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            int errorId = 0;
            int col = 1;
            int row = 1;
            var header = string.Empty;

            if (reportParameters.IsCurrentState || reportParameters.IsAllState)
            {
                var makerRkk = reportParameters.IsObjectsByAuthor ? "исполнителя " + reportParameters.Author[Guids.ГруппыИПользователи.Поля.КороткоеИмя].ToString() :
                    reportParameters.IsObjectsByDepartment ? "подразделения " + reportParameters.Department.ToString() :
                    reportParameters.Chiefs ? "всех объектов у руководителей" : "всех объектов";
                var beginHeader = reportParameters.IsCurrentState ? "Текущие состояния" : "Состояния";
                header = beginHeader + " процесса " + reportParameters.ProcRefObj + ". Для " + makerRkk;
            }
            if (reportParameters.IsAllState)
            {
                header += " c " + reportParameters.BeginPeriod.ToShortDateString() + " по " + reportParameters.EndPeriod.ToShortDateString();
            }

            if (reportParameters.IsSelectedState)
            {
                header = "Статистика по исполнителям процесса для этапа " + reportParameters.StatesObject.First().ToString() +
                    " c " + reportParameters.BeginPeriod.ToShortDateString() + " по " + reportParameters.EndPeriod.ToShortDateString();
            }

            PasteHeadName(xls, col, row, 12, 5, header);
            row++;

            col = 1;
            row = PasteHeader(xls, reportParameters, col, row);
            var beginRow = row;
            var sortingReportData = SortDataBySettings(reportParameters, reportData);
            var rowwithDublicateStartedObject = row;
            var prewMaker = string.Empty;
            var prewStartedObject = string.Empty;
            var objectsForSum = new List<ActionParameters>();

            foreach (var obj in sortingReportData)
            {
                try
                {
                    if (obj != null)
                    {
                        progressBar.PerformStep();

                        if (prewMaker != string.Empty && prewMaker != obj.MakerName.ToString())
                        {
                            row = SummaryWriteRow(xls, reportParameters, logDelegate, objectsForSum, row, prewMaker);
                            objectsForSum = new List<ActionParameters>();
                            beginRow = row;
                        }
                        objectsForSum.Add(obj);
                        //Заполнение параметров объекта Номенклатуры
                        var startedObject = obj.RKK.ToString();
                        row = WriteRow(xls, reportParameters, obj, logDelegate, row, startedObject);
                        MergeRegNumber(xls, reportParameters, row, ref rowwithDublicateStartedObject, ref prewStartedObject, startedObject);
                        prewMaker = obj.MakerName.ToString();
                    }
                    errorId = 6;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("4. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки  " + errorId.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            MergeRegNumber(xls, reportParameters, row + 1, ref rowwithDublicateStartedObject, ref prewStartedObject, null);
            if (reportParameters.IsSelectedState)
            {
                SummaryWriteRow(xls, reportParameters, logDelegate,  objectsForSum, row, prewMaker);
            }
            PrintSetting(xls, reportParameters);
            xls.AutoHeight();

            if (reportParameters.IsSelectedState)
            {
                xls.AddWorksheet("Диаграммы", true);
                row = 1;
                PasteHeadName(xls, col, row, 12, 8, "Статистика по исполнителям процесса для этапа " + reportParameters.StatesObject.First() +
                   " c " + reportParameters.BeginPeriod.ToShortDateString() + " по " + reportParameters.EndPeriod.ToShortDateString());

                GetCharts(xls, sortingReportData);
                xls.SelectWorksheet("Лист1");
                xls.Worksheet.Name = "Статистика по исполнителям";
            }
        }

        private static void MergeRegNumber(Xls xls, ReportParameters reportParameters, int row, ref int rowwithDublicateStartedObject, ref string prewStartedObject, string startedObject = null)
        {
            if ((reportParameters.IsCurrentState || reportParameters.IsAllState) && (startedObject == null || prewStartedObject != startedObject))
            {
                var mergeCount = row - rowwithDublicateStartedObject;
                if (mergeCount > 1)
                {
                    xls[1, rowwithDublicateStartedObject - 1, 1, mergeCount].MergeCells = true;
                }
                prewStartedObject = startedObject;
                rowwithDublicateStartedObject = row;
            }
        }

        private static List<ActionParameters> SortDataBySettings(ReportParameters reportParameters, List<ActionParameters> reportData)
        {
            var sortingReportData = new List<ActionParameters>();

            if (reportParameters.IsTypeOrder && reportParameters.IsCurrentState)
            {
                sortingReportData.AddRange(reportData.OrderBy(t => t.RefObject.ToString()).ThenBy(t => t.RKK[Guids.Ркк.Поля.Заказчик].ToString()).ThenBy(t => t.RKK[Guids.Ркк.Поля.ШифрЗаказа].ToString()));
            }
            else
            {
                if (reportParameters.IsCurrentState)
                {
                    if (reportParameters.Chiefs)
                    {
                        sortingReportData.AddRange(reportData.OrderBy(t => t.MakerName).ThenBy(t => t.RKK.ToString()).ThenBy(t => t.StartTime));
                    }
                    else
                    {
                        sortingReportData.AddRange(reportData.OrderBy(t => t.RKK.SystemFields.CreationDate).ThenBy(t => t.RKK.ToString()).ThenBy(t => t.MakerName).ThenBy(t => t.StartTime));
                    }
                }
                else
                {
                    if (reportParameters.IsAllState)
                    {
                        sortingReportData.AddRange(reportData.OrderBy(t => t.RKK[Guids.Ркк.Поля.РегистрационныйНомер].ToString()).ThenBy(t => t.RKK[Guids.Ркк.Поля.От].ToString()).ThenBy(t => t.RefObject.SystemFields.CreationDate));
                    }
                    else
                    {
                        //if (reportParameters.IsSelectedState)
                        //{
                        //    sortingReportData.AddRange(reportData.Where(t => t.RKK[Guids.Ркк.Поля.РегистрационныйНомер].ToString().Contains("/914")).OrderBy(t => t.MakerName.ToString()).ThenBy(t => t.StartTime).ThenBy(t => t.EndTime));
                        //}
                        //else
                        {
                            sortingReportData.AddRange(reportData.OrderBy(t => t.MakerName.ToString()).ThenBy(t => t.StartTime).ThenBy(t => t.EndTime));
                        }
                    }
                }
            }

            return sortingReportData;
        }

        public int PasteHeadName(Xls xls, int col, int row, int sizeFont, int countMerge, string head)
        {
            xls[1, row].SetValue(head);
            xls[1, row].Font.Bold = true;
            xls[col, row].Font.Size = sizeFont;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, countMerge, 1].MergeCells = true;
            row++;

            return row;
        }

        private static void PrintSetting(Xls xls, ReportParameters reportParameters)
        {
            var col = 1;
            xls.SetColumnWidth(col, 16);
            col++;
            if (reportParameters.IsCurrentState || reportParameters.IsAllState)
            {
                xls.SetColumnWidth(col, 36);
                col++;
            }
            xls.SetColumnWidth(col, 20);
            col++;
            xls.SetColumnWidth(col, 36);
            col++;
            xls.SetColumnWidth(col, 16);
            col++;
            xls.SetColumnWidth(col, 25);   
        }

        public string TermToString(TimeSpan term)
        {
            string termToString = string.Empty;

            if (term.Hours > 0 || term.Minutes > 0 || term.Seconds > 0)
            {
                termToString += (term.Days*24  + term.Hours).ToString() + ":" + term.Minutes + ":" + term.Seconds;
            }

            return termToString;
        }

        public string SecondToTime (double totalSecond)
        {
            var hours = Math.Ceiling(TimeSpan.FromSeconds(totalSecond).TotalHours);
            var minutes = TimeSpan.FromSeconds(totalSecond).Minutes;
            var seconds = TimeSpan.FromSeconds(totalSecond).Seconds;      

            return hours + ":" + minutes + ":" + seconds;
        }

        public int GetCharts(Xls xlsDiagram, List<ActionParameters> reportData)
        {
            var row = 2;
            //Добавление графика
            Worksheet sheet = xlsDiagram.Worksheet;
            Dictionary<string, List<ActionParameters>> makerActions = new Dictionary<string, List<ActionParameters>>();

            var makers = reportData.Select(t => t.MakerName).Distinct().ToList();

            foreach (var maker in makers)
            {
                makerActions.Add(maker, reportData.Where(t => t.MakerName == maker).ToList());
            } 
            xlsDiagram[1, row].SetValue("Исполнитель");
            xlsDiagram[2, row].SetValue("Кол-во заданий");
            xlsDiagram[3, row].SetValue("Общее время выполнения");
            xlsDiagram[4, row].SetValue("Среднее время выполнения");

            xlsDiagram[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                        XlBordersIndex.xlEdgeBottom,
                                                                                        XlBordersIndex.xlEdgeLeft,
                                                                                        XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            xlsDiagram[1, row, 4, 1].Font.Bold = true;
            xlsDiagram[1, row, 4, 1].Font.Size = 11;
            xlsDiagram[1, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            row++;
            var beginRow = row;

            foreach (var makerAct in makerActions.OrderByDescending(t=>t.Value.Count))
            {
                xlsDiagram[1, row].SetValue(makerAct.Key);
                xlsDiagram[2, row].SetValue(makerAct.Value.Count );
                //  xlsDiagram[3, row].SetValue(SecondToTime(makerAct.Value.Where(t=> t.EndTime != DateTime.MinValue)
                //     .Sum(t=>Convert.ToInt32((t.EndTime - t.StartTime).TotalSeconds)/makerAct.Value.Count)));
                xlsDiagram[3, row].SetValue(SecondToTime(makerAct.Value.Sum(t => ПолучитьЗадержкуВЧасах(t.StartTime, t.EndTime).TotalSeconds)));
                xlsDiagram[4, row].NumberFormat = @"hh:mm:ss";
                xlsDiagram[4, row].SetFormula("= " + xlsDiagram[3, row].Address + "/" + xlsDiagram[2, row].Address);

                xlsDiagram[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                            XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                row++;
            }

            var height = Convert.ToInt32(Math.Round((row+1) * 14.5));

            xlsDiagram.AutoWidth();

            ChartObjects xlCharts = (ChartObjects)sheet.ChartObjects();
            ChartObject myChart = (ChartObject)xlCharts.Add(0, height, 500, height * 1.4);
            Chart chart = myChart.Chart;
            SeriesCollection seriesCollection = (SeriesCollection)chart.SeriesCollection();

            Series series = seriesCollection.NewSeries();
            series.Name = "Статистика по количеству заданий";

            series.XValues = sheet.get_Range(xlsDiagram[1, beginRow].Address, xlsDiagram[1,row-1].Address);
            series.Values = sheet.get_Range(xlsDiagram[2, beginRow].Address, xlsDiagram[2, row-1].Address);          

            chart.ChartType = XlChartType.xlBarClustered;
            chart.ClearToMatchStyle();
            chart.ChartStyle = 218;
            chart.Legend.Delete();

            var indx = 1;

            foreach (var makerAct in makerActions.OrderByDescending(t => t.Value.Count))
            {
                chart.SeriesCollection(1).DataLabels(indx).Format.TextFrame2.TextRange = makerAct.Value.Count + " (" + Math.Round(Convert.ToDouble(makerAct.Value.Count) * 100 / Convert.ToDouble(reportData.Count),1) + " %)";
                indx++;
            }

            foreach (Microsoft.Office.Interop.Excel.Point point in series.Points())
            {
                point.Format.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                point.Format.Fill.Transparency = 0;
                point.Format.Fill.Solid();

            }
            DataLabels dl = series.DataLabels();
            dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0;

            ChartObjects xlCharts2 = (ChartObjects)sheet.ChartObjects();
            ChartObject myChart2 = (ChartObject)xlCharts2.Add(0, height * 2.4, 500, height * 1.4);
            Chart chart2 = myChart2.Chart;
            SeriesCollection seriesCollection2 = (SeriesCollection)chart2.SeriesCollection();

            Series series2 = seriesCollection2.NewSeries();
            series2.Name = "Статистика по среднему времени";

            series2.XValues = sheet.get_Range(xlsDiagram[1, beginRow].Address, xlsDiagram[1, row-1].Address);
            series2.Values = sheet.get_Range(xlsDiagram[4, beginRow].Address, xlsDiagram[4, row-1].Address);

            chart2.ChartType = XlChartType.xlBarClustered;
            chart2.ClearToMatchStyle();
            chart2.ChartStyle = 218;
            chart2.Legend.Delete();


            foreach (Microsoft.Office.Interop.Excel.Point point in series2.Points())
            {
                point.Format.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue;
                point.Format.Fill.Transparency = 0;
                point.Format.Fill.Solid();
            }
            DataLabels dl2 = series2.DataLabels();
            dl2.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0;

            return row;
        }

        private int WriteRow(Xls xls, ReportParameters reportParametetrs, ActionParameters action, LogDelegate logDelegate, int row, string numberRkk)
        {
            logDelegate(String.Format("Запись объекта: " + numberRkk));
            int prewRow = row;
            var col = 1;
            xls[col, row].SetValue(numberRkk);
            col++;

            if (reportParametetrs.IsCurrentState || reportParametetrs.IsAllState)
            {
                xls[col, row].SetValue(action.RefObject.ToString());
                col++;
            }

            xls[col, row].SetValue(action.MakerName);
            col++;

            if (action.EndTime != DateTime.MinValue)
            {
                xls[col, row].SetValue(action.StartTime + " - " + action.EndTime);
                col++;
                if (!reportParametetrs.Chiefs)
                {
                    xls[col, row].SetValue((reportParametetrs.IsTypeOrder && reportParametetrs.IsCurrentState) ? action.RKK[Guids.Ркк.Поля.Заказчик].ToString() :
                        reportParametetrs.IsControlState ? ПолучитьЗадержкуВДнях(action.StartTime, action.EndTime, reportParametetrs.CountControlDays).ToString() :
                        reportParametetrs.IsSelectedState ? ПолучитьЗадержкуВЧасах(action.StartTime, action.EndTime).ToString() : (action.EndTime - action.StartTime).ToString());
                    col++;
                }
            }
            else
            {
                xls[col, row].SetValue(action.StartTime + " - ");
                col++;
                if (!reportParametetrs.Chiefs)
                {
                    xls[col, row].SetValue((reportParametetrs.IsTypeOrder && reportParametetrs.IsCurrentState) ? action.RKK[Guids.Ркк.Поля.Заказчик].ToString() :
                         reportParametetrs.IsControlState ? ПолучитьЗадержкуВДнях(action.StartTime, action.EndTime, reportParametetrs.CountControlDays).ToString() :
                         reportParametetrs.IsSelectedState ? ПолучитьЗадержкуВЧасах(action.StartTime, action.EndTime).ToString() : " - ");
                    col++;
                }
            }
            xls[col, row].SetValue((reportParametetrs.IsTypeOrder && reportParametetrs.IsCurrentState) ? action.RKK[Guids.Ркк.Поля.ШифрЗаказа].ToString() : action.Comment);

            if (reportParametetrs.Chiefs)
            {
                if (action.RKK.Class.IsInherit(Guids.Ркк.Типы.Зпзп))
                {
                    xls[col, row].SetValue(action.RKK[Guids.Ркк.Поля.СодержаниеЗпзп].GetString());
                    col++;
                    xls[col, row].SetValue(action.RKK[Guids.Ркк.Поля.КонтрагентЗпзп].GetString());
                    col++;
                    xls[col, row].SetValue(action.RKK[Guids.Ркк.Поля.СуммаЗпзп].GetString());
                }
                else
                {
                    if (action.RKK.Class.IsInherit(Guids.Ркк.Типы.Сзно))
                    {
                        xls[col, row].SetValue(action.RKK[Guids.Ркк.Поля.СодержаниеСзно].GetString());
                        col++;
                        xls[col, row].SetValue(action.RKK[Guids.Ркк.Поля.КонтрагентСзно].GetString());
                        col++;
                        xls[col, row].SetValue(action.RKK[Guids.Ркк.Поля.СуммаСзно].GetString());
                    }
                }
            }

            xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                               XlBordersIndex.xlEdgeBottom,
                                                                                               XlBordersIndex.xlEdgeLeft,
                                                                                               XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            row++;
            xls[1, prewRow, col, row - prewRow].WrapText = true;
            xls[1, prewRow, col, row - prewRow].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, prewRow, col, row - prewRow].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            return row;
        }

        private int SummaryWriteRow(Xls xls, ReportParameters reportParameters, LogDelegate logDelegate, List<ActionParameters> objects, int row, string maker)
        {       
            var col = 1;
            xls[col, row].SetValue("Итого: " + objects.Count);
            col++;
            xls[col, row].SetValue("Среднее время на этапе у " + maker);

            if (reportParameters.IsSelectedState)
            {
                logDelegate("Сумма для " + maker);
                xls[col, row, 2, 1].Merge();
                col++;
                col++;
                xls[col, row].SetValue(SecondToTime(objects.Sum(t => ПолучитьЗадержкуВЧасах(t.StartTime, t.EndTime).TotalSeconds)));
                col++;
                xls[col, row].NumberFormat = @"hh:mm:ss";
                xls[col, row].SetFormula("= " + xls[col - 1, row].Address + "/" + objects.Count);
                col++;
                row++;

                xls[1, row - 1, col, 1].Font.Bold = true;
                xls[1, row - 1, col, 1].Font.Size = 11;
                xls[1, row - 1, col, 1].HorizontalAlignment = XlHAlign.xlHAlignRight;

                xls[1, row - 1 , col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                   XlBordersIndex.xlEdgeBottom,
                                                                                                   XlBordersIndex.xlEdgeLeft,
                                                                                                   XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            }
            return row;
        }
    }
}
