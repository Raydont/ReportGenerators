using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.ActiveAction;
using TFlex.DOCs.Model.References.Processes;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace TransitTimeProcess
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
                selectionForm.Init(context);

                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    var reportClass = new ExceedingTermsReport();
                    reportClass.MakeReport(selectionForm.reportParameters);
                }

                

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK,
                                                     System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
            }
        }
    }

    public delegate void LogDelegate(string line);
    public class ExceedingTermsReport
    {
        public IndicationForm m_form = new IndicationForm();
        public void Dispose()
        {
        }

        public void MakeReport(ReportParameters reportParameters)
        {
            ExceedingTermsReport report = new ExceedingTermsReport();
            LogDelegate m_writeToLog;

            report.Make(reportParameters);
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

            Xls xls = new Xls();
            m_form.setStage(IndicationForm.Stages.DataAcquisition);
            Dictionary<ReferenceObject, List<MakerComment>> reportData = new Dictionary<ReferenceObject, List<MakerComment>>();

            if (reportParameters.SelectedObject)
            {
                reportData.Add(reportParameters.RefObject,ReadData(reportParameters.RefObject, m_writeToLog, m_form.progressBar));
            }
            else
            {
                var referenceInfo = ReferenceCatalog.FindReference(reportParameters.RefObject.Reference.Id);
                var reference = referenceInfo.CreateReference();
                

                var filter = new TFlex.DOCs.Model.Search.Filter(referenceInfo);
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.CreationDate], ComparisonOperator.GreaterThan, reportParameters.BeginPeriod);
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.CreationDate], ComparisonOperator.LessThan, reportParameters.EndPeriod);
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, reportParameters.RefObject.Class);
                var findedObjects = reference.Find(filter).ToList();

                MessageBox.Show(string.Join("\r\n", findedObjects.Select(t=>t.ToString())));

                foreach (var obj in findedObjects)
                {
                    var makerComments = ReadData(obj, m_writeToLog, m_form.progressBar);
                    if (makerComments != null)
                        reportData.Add(obj, makerComments);
                }

                MessageBox.Show(2 + "");

            }
            if (reportData.Count == 0)
            {
                this.Dispose();
                m_form.Close();
                return;
            }
                

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
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                xls.Application.DisplayAlerts = true;
                xls.Visible = true;
            }
            m_form.Dispose();
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

        public int PasteHeader(Xls xls, int col, int row)
        {
            col = InsertHeader("Наименование этапа", col, row, xls);
            col = InsertHeader("Время на этапе", col, row, xls);        
            col = InsertHeader("Фактическое время на этапе", col, row, xls);
            col = InsertHeader("Срок", col, row, xls);
            col = InsertHeader("Комментарий", col, row, xls);
            col = InsertHeader("Исполнитель", col, row, xls);

            xls[1, row, 6, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            row++;
            return row;
        }


        public List<MakerComment> ReadData(ReferenceObject nomObject, LogDelegate logDelegate, ProgressBar progressBar)
        {
            List<MakerComment> stateActionsObjects = new List<MakerComment>();

            //Последний процесс текущего объекта по дате создания
            var lastProcess = ProcessesReference.GetProcessesByObject(nomObject).OfType<WorkflowReferenceObject>().ToList().
                     Where(t => !t.ToString().Contains("Доработка")
                             && t.ToString() != "Добавление параметров организации"
                             && t.ToString() != "Установка подписи"
                             && !t.ToString().Contains("Доступ")
                             && !t.ToString().Contains("Изменение параметра ДЗУ")
                             && t.Links.ToMany[Guids.LinkProcessToAction].Objects.Count(s => s.ToString().Contains("Документ уже запущен по маршруту")) == 0).
                             OrderByDescending(t => t.SystemFields.CreationDate).FirstOrDefault();

            //  var st = lastProcess + string.Join("\r\n", lastProcess.Links.ToMany[Guids.LinkProcessToAction].Objects.Select(t => t.ToString()));
            //   MessageBox.Show(st);


            if (lastProcess == null)
                return null;

            if (Convert.ToInt32(lastProcess[Guids.ProcessStatusGuid].Value) == 3)
            {
              //  MessageBox.Show("Процесс для объекта " + nomObject.ToString() + " ПРЕРВАН! Обратитесь в отдел 911 для перезапуска процесса", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            ProcessReferenceObject process = lastProcess as ProcessReferenceObject;
            ProcessReferenceObject parentProcess = null;
            if (process != null)
            {
                // Текущеее действие родительского процесса
                ActiveActionReferenceObject parentActiveAction = process.ParentActiveAction;

                if (parentActiveAction != null)
                {
                    // Родительский процесс
                    parentProcess = parentActiveAction.Process;
                }             
            }
            //действия процесса
            var activeActions = lastProcess.GetObjects(Guids.ActiveActionsLinkGuid);
            //Действия процесса
            var actions = lastProcess.Links.ToMany[Guids.LinkProcessToAction].Objects.ToList();
            //Последнее действие процесса
            var lastAction = actions.Last();
            //Все состояния
            var states = lastProcess.Procedure.GetObjects(Guids.ListObjectsSate).ToList();

            if (parentProcess != null)
            {
                activeActions.AddRange(parentProcess.GetObjects(Guids.ActiveActionsLinkGuid));
                actions.AddRange(parentProcess.Links.ToMany[Guids.LinkProcessToAction].Objects);
                states.AddRange(parentProcess.Procedure.GetObjects(Guids.ListObjectsSate).ToList());
            }
            //Текущие действия процесса
            var currentActions = actions.Where(t => t.Class.IsInherit(Guids.StateActionType)).
                Where(t => activeActions.Select(a => ((Guid)a[Guids.IdentificatorStepGuid].Value)).
                Contains((Guid)t[Guids.IdenticatorState].Value)).ToList();
            //   if (currentActions.Count() > 0 && !currentActions[0].ToString().ToLower().Contains("задержка"))


            stateActionsObjects.AddRange(GetAllMakerComment(actions, states));

            //string statesStr = string.Join("\r\n", listMakerComment.Select(t => t.ToString()));
            //Clipboard.SetText(statesStr);
            //MessageBox.Show(statesStr);


            return stateActionsObjects;
        }

        public static List<MakerComment> GetAllMakerComment(List<ReferenceObject> actions, List<ReferenceObject> states)
        {
            List<MakerComment> listMakerComment = new List<MakerComment>();
            int errorId = 0;

            ////Выбор состояний типа Работа или Согласование
            //var statesByType = states.Where(s => s.Class.IsInherit(ClassGuids.TypeStateWorkGuid) || s.Class.IsInherit(ClassGuids.TypeStateCoordinationGuid)).ToList();
            //Действия работа или согласование
            try
            {

                var workActions = actions.Where(t => t.Class.IsInherit(Guids.StateActionType) && states.Where(s => s.Class.IsInherit(Guids.TypeStateWorkGuid)
                                                                   || s.Class.IsInherit(Guids.TypeStateCoordinationGuid)).
                                                                               Select(s => s.SystemFields.Guid).
                                                Contains((Guid)t[Guids.IdenticatorState].Value)).ToList();
                errorId = 1;
                if (workActions.Count() > 0)
                {
                    errorId = 2;
                    var statesWork = states.Where(t => workActions.Select(a => ((Guid)a[Guids.IdenticatorState].Value)).
                                                       Contains((t.SystemFields.Guid))).
                                                       Where(t => t.Class.IsInherit(Guids.TypeStateWorkGuid) || t.Class.IsInherit(Guids.TypeStateCoordinationGuid)).ToList();
                    errorId = 3;
                    //Добавление в список найденных пользователей и их комментариев
                    listMakerComment.AddRange(GetListMakerComment(actions, statesWork, workActions, 1));
                    errorId = 4;
                }
            }
            catch
            {
                MessageBox.Show("Некорректно считались данные у состояний\r\n " + string.Join("\r\n", states.Select(t => t.ToString())) + "\r\nкод ошибки 2." + errorId, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return listMakerComment;
        }

        //Добавление в список найденных пользователей и их комментариев
        private static List<MakerComment> GetListMakerComment(List<ReferenceObject> actions, List<ReferenceObject> statesByType, List<ReferenceObject> searchedActions, byte workType)
        {
            Dictionary<ReferenceObject, List<ReferenceObject>> stateActionsByType = new Dictionary<ReferenceObject, List<ReferenceObject>>();
            foreach (var state in statesByType)
            {
                stateActionsByType.Add(state, searchedActions.Where(t => t.Class.IsInherit(Guids.StateActionType)).Where(t => ((Guid)t[Guids.IdenticatorState].Value) == state.SystemFields.Guid).ToList());
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
                listActions.AddRange(GetActionToWorkState(action.Parents.Where(t => !statesByType.Select(s => s.SystemFields.Guid).Contains((Guid)t[Guids.IdenticatorState].Value)).ToList(), statesByType));
            }
            return listActions;
        }



        private void MakeSelectionByTypeReport(Xls xls, ReportParameters reportParameters, Dictionary<ReferenceObject, List<MakerComment>> reportData, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            var progressStep = Convert.ToDouble(progressBar.Maximum) / (Convert.ToDouble(reportData.Where(t => t.Value != null).Sum(t => t.Value.Count)));

            if (progressBar.Maximum < (Convert.ToDouble(reportData.Where(t => t.Value != null).Sum(t => t.Value.Count))))
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }


            double step = 0;

            int errorId = 0;
            int col = 1;
            int row = 1;
            foreach (var obj in reportData)
            {
                PasteHeadName(xls, col, row, 12, "Сроки прохождения состояний процесса для объекта " + obj.Key.ToString());
                row++;
                col = 1;
                row = PasteHeader(xls, col, row);
                var rowMain = row;
                try
                {
                    errorId = 1;
                    if (obj.Value != null)
                    {
                        foreach (var stateAct in obj.Value.OrderBy(t => t.Actions.First().StartTime))
                        {
                            progressBar.PerformStep();
                            //Заполнение параметров объекта Номенклатуры
                            row = WriteRow(xls, stateAct, logDelegate, row);

                        }
                    }
                    errorId = 6;
                    PrintSetting(xls);
                    xls[1, 1, 6, row].WrapText = true;
                    xls[1, rowMain, 6, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[1, rowMain, 6, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    row++;

                }
                catch (Exception ex)
                {
                    MessageBox.Show("4. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки  " + errorId.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }

      
        public int PasteHeadName(Xls xls, int col, int row, int sizeFont, string head)
        {
            xls[1, row].SetValue(head);
            xls[1, row].Font.Bold = true;
            xls[col, row].Font.Size = sizeFont;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 6, 1].MergeCells = true;

            row++;

            return row;
        }

        private static void PrintSetting(Xls xls)
        {
            xls.SetColumnWidth(1, 47);
            xls.SetColumnWidth(2, 36);
            xls.SetColumnWidth(3, 15);
            xls.SetColumnWidth(4, 11);
            xls.SetColumnWidth(5, 22);
            xls.SetColumnWidth(6, 20);
        }

        private int WriteRow(Xls xls, MakerComment stateAction, LogDelegate logDelegate, int row)
        {
            logDelegate(String.Format("Запись объекта номенклатуры: " + stateAction.CurrentState));
            int prewRow = row;

            xls[1, row].SetValue(stateAction.CurrentState);          
            xls[4, row].SetValue(stateAction.TermToString(stateAction.DurationState).ToString());

            foreach (var action in stateAction.Actions)
            {
                if (action.EndTime != DateTime.MinValue)
                {
                    xls[2, row].SetValue(action.StartTime + " - " + action.EndTime);
                    xls[3, row].SetValue(stateAction.TermToString(action.EndTime - action.StartTime).ToString());
                }
                else
                {
                    xls[2, row].SetValue(action.StartTime + " - ");
                    xls[3, row].SetValue(" - ");
                }

                xls[5, row].SetValue(action.Comment);
                xls[6, row].SetValue(action.MakerName);

                if ((action.EndTime - action.StartTime) > stateAction.DurationState || (action.EndTime == DateTime.MinValue && (DateTime.Now - action.StartTime) > stateAction.DurationState))
                {
                    xls[2, row, 2, 1].Interior.Color = Color.OrangeRed;
                    xls[5, row, 2, 1].Interior.Color = Color.OrangeRed;
                }

                xls[1, row, 6, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                   XlBordersIndex.xlEdgeBottom,
                                                                                                   XlBordersIndex.xlEdgeLeft,
                                                                                                   XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                row++;
            }
            if (stateAction.Actions.Count == 0)
                row++;
            else
            {
                if (stateAction.Actions.Count > 1)
                {
                    xls[1, prewRow, 1, (row-prewRow)].MergeCells = true;
                    xls[4, prewRow, 1, (row-prewRow)].MergeCells = true;
                }
            }

            return row;
        }   
    }
}
