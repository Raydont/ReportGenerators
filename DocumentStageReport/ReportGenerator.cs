using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using System.IO;
using ReportHelpers;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Filter = TFlex.DOCs.Model.Search.Filter;
using Globus.DOCs.Technology.Reports.ServerSide.Dialogs;
using System.Reflection;
using System.Text;
using System.Drawing;
using TFlex.DOCs.Model.Macros.ObjectModel.Layout;

namespace DocumentStagesReport
{
    public class ReportGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xlsx";
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
            var reportForm = new ReportForm();
            using (new WaitCursorHelper(false))
            {
                if (reportForm.ShowDialog() == DialogResult.OK)
                {
                    // Даты начала и окончания
                    DateTime beginDate = reportForm.beginDate; beginDate = beginDate.Date;
                    DateTime endDate = reportForm.endDate; endDate = endDate.Date.AddDays(1).AddTicks(-1);
                    // Тип выбранного объекта
                    ClassObject objectClass = GetObjectClass(context.ObjectsInfo[0].ObjectId);
                    // Чтение данных для отчета
                    var reportTable = ReadReportData(beginDate, endDate, objectClass);

                    if (reportTable.Count == 0)
                    {
                        MessageBox.Show("За период не найдено ни одного активного документа");
                        return;
                    }

                    bool isCatch = false;
                    if (File.Exists(context.ReportFilePath))
                    {
                        try
                        {
                            File.Delete(context.ReportFilePath);
                        }
                        catch
                        {
                            MessageBox.Show("Закройте файл " + context.ReportFilePath + " перед формированием отчета!");
                            return;
                        }
                    }

                    context.CopyTemplateFile();    // Создаем копию шаблона   

                    // Открытие нового документа
                    Xls xls = new Xls();
                    try
                    {
                        xls.Application.DisplayAlerts = false;
                        xls.Open(context.ReportFilePath);

                        // Создание заголовка
                        CreateHeaderV1(reportTable.Select(t => t.Value.Count).Max(), beginDate, endDate, objectClass, xls);
                        //CreateHeader(reportTable.Select(t => t.Value.Count).Max(), beginDate, endDate, objectClass, xls);
                        // Создание таблицы отчета
                        CreateTableV1(5, reportTable, xls);
                        //CreateTable(6, reportTable, xls);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Исключительная ситуация!", ex.Message);
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
                            Process.Start(context.ReportFilePath);
                        }
                    }
                }
            }
        }
        // Класс объекта в журнале регистрации
        private ClassObject GetObjectClass(int objectID)
        {
            // Справочник РКК
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.RegCardsReference.RefGuid);
            var reference = referenceInfo.CreateReference();
            //Поиск объекта по его ID, затем возврат его класса
            var filter = new Filter(referenceInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, objectID);
            return reference.Find(filter).FirstOrDefault().Class;
        }         

        private Dictionary<string, List<ReportString>> ReadReportData(DateTime beginDate, DateTime endDate, ClassObject classObject)
        {
            // Справочник РКК
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.RegCardsReference.RefGuid);
            var reference = referenceInfo.CreateReference();

            // Все объекты указанного класса с датой регистрации внутри указанного периода, со статусом процесса - Выполняется (1)
            var filter = new Filter(referenceInfo);
            filter.Terms.AddTerm("[" + Guids.RegCardsReference.Fields.RegDate + "]", ComparisonOperator.GreaterThanOrEqual, beginDate);
            filter.Terms.AddTerm("[" + Guids.RegCardsReference.Fields.RegDate + "]", ComparisonOperator.LessThanOrEqual, endDate);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);

            var reportTable = new Dictionary<string, List<ReportString>>();

            var foundObjects = reference.Find(filter);
            if (foundObjects.Count == 0) return reportTable;

            // Список процедур, по которым будет идти отбор процессов
            List<Guid> proceduresList = new List<Guid> { Guids.ProcedureReference.Object.ApprovalDZU,
                                                         Guids.ProcedureReference.Object.ApprovalSZ,
                                                         Guids.ProcedureReference.Object.ApprovalSZNO,
                                                         Guids.ProcedureReference.Object.ApprovalSZNPDS,
                                                         Guids.ProcedureReference.Object.ApprovalZICD,
                                                         Guids.ProcedureReference.Object.ApprovalZPZP,
                                                         Guids.ProcedureReference.Object.ProductOrder,
                                                         Guids.ProcedureReference.Object.TaxesSZNO };

            // Поиск активных процессов по объектам, отобранным за период
            referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.ProcessReference.RefGuid);
            reference = referenceInfo.CreateReference();
            filter = new Filter(referenceInfo);
            filter.Terms.AddTerm("[" + Guids.ProcessReference.Fields.ProcessStartObjects + "]", ComparisonOperator.IsOneOf, foundObjects.ToArray());
            filter.Terms.AddTerm("[" + Guids.ProcessReference.Fields.ProcessStatus + "]", ComparisonOperator.Equal, 1);
            filter.Terms.AddTerm("[" + Guids.ProcessReference.Fields.LinkProcessToProcedure + "]->[Guid]", ComparisonOperator.IsOneOf, proceduresList.ToArray());
            var foundProcesses = reference.Find(filter);
            if (foundProcesses.Count == 0) return reportTable;

            List<Guid> fictiveUsers = new List<Guid> { Guids.UsersReference.Objects.DeadlineControl };

            var progress = new ProgressDialog();
            progress.StartPosition = FormStartPosition.CenterScreen;
            progress.Show();
            progress.Text = "Формирование отчета";
            progress.SetInfo("Чтение данных из журнала РКК");
            var offset = 0;
            var count = foundObjects.Count;

            try
            {
                foreach (var process in foundProcesses)
                {
                    var rkkObject = process.GetObjects(Guids.ProcessReference.Fields.ProcessStartObjects).FirstOrDefault();
                    offset++;
                    progress.SetInfo("Чтение данных документа: " + rkkObject.ToString());
                    progress.SetProgress(offset, count);

                    var reportList = new List<ReportString>();
                    // Действия процесса (только пользовательские)
                    var actions = process.GetObjects(Guids.ProcessReference.Fields.LinkProcessToAction).
                        Where(t => t.Class.Guid == Guids.ProcessActionsReference.Types.UserStateActionType).ToList();
                    foreach (var action in actions)
                    {
                        ReportString reportString = new ReportString();
                        // Этап
                        string stage = action.ToString().Replace("Состояние: ", "");
                        stage = stage.IndexOf(":") == -1 ? stage : stage.Substring(0, stage.IndexOf(":"));
                        reportString.Stage = stage;
                        // Время начала
                        reportString.StartTime = action[Guids.ProcessActionsReference.Fields.StartTime].Value.ToString();
                        // Время окончания
                        reportString.EndTime = action[Guids.ProcessActionsReference.Fields.EndTime].Value == null ? String.Empty :
                            action[Guids.ProcessActionsReference.Fields.EndTime].Value.ToString();
                        var startTime = (DateTime)action[Guids.ProcessActionsReference.Fields.StartTime].Value;
                        var endTime = action[Guids.ProcessActionsReference.Fields.EndTime].Value == null ? DateTime.Now :
                            (DateTime)action[Guids.ProcessActionsReference.Fields.EndTime].Value;
                        var diff = endTime - startTime;
                        // Время выполнения этапа
                        reportString.StageInterval = String.Format("{0}.{1}:{2}:{3}", diff.Days, diff.Hours, diff.Minutes, diff.Seconds);
                        // Текущее время (для сортироки)
                        reportString.CurrentTime = endTime;
                        // Исполнители
                        var executors = action.GetObjects(Guids.ProcessActionsReference.Fields.MakersProcessActions);
                        if (executors.Count > 0)
                            reportString.Executors.AddRange(executors.Select(t => t.ToString()));
                        // Поиск запрещенного юзера
                        var user = executors.Select(t => t.GetObject(Guids.ProcessActionsReference.Fields.LinkMakerToUser)).FirstOrDefault(t => fictiveUsers.Contains(t.SystemFields.Guid));
                        // Если юзер - не запрещенный, то  записываем в коллекцию этапов
                        if (user == null)
                            reportList.Add(reportString);
                    }
                    reportTable.Add(rkkObject.ToString() + " от " + ((DateTime)rkkObject[Guids.RegCardsReference.Fields.RegDate]).ToShortDateString(), 
                        reportList.OrderBy(t => t.CurrentTime).ToList());
                }
            }
            catch (Exception ex)
            {
                progress.Close();
                MessageBox.Show(ex.ToString(), "Исключительная ситуация!");
            }
            progress.Close();
            return reportTable.OrderBy(t => t.Key).ToDictionary(t => t.Key, t => t.Value);
        }
        // Заголовок отчета + шапка таблицы
        private void CreateHeader(int stagesCount, DateTime beginDate, DateTime endDate, ClassObject classObject, Xls xls)
        {
            // Заголовок отчета
            xls[2, 1].SetValue(String.Format("Состояния объектов РКК \"{0}\" за период с {1} по {2}", classObject, beginDate.ToShortDateString(), endDate.ToShortDateString()));
            xls[2, 1].Font.Bold = true;
            xls[2, 1].Font.Size = 14;

            // Шапка таблицы
            xls[1, 3].SetValue("Документ");
            xls[1, 3].ColumnWidth = 35;
            xls[1, 3, 1, 3].Merge();
            for (int i = 0; i < stagesCount; i++)
            {
                xls[2 + i * 3, 3].SetValue("Этап " + (i + 1));
                xls[2 + i * 3, 3, 3, 1].Merge();
                xls[2 + i * 3, 4].SetValue("Исполнители");
                xls[2 + i * 3, 4, 3, 1].Merge();
                xls[2 + i * 3, 4].SetValue("Время начала");
                xls[3 + i * 3, 4].SetValue("Время завершения");
                xls[4 + i * 3, 4].SetValue("Время этапа");
                xls[2 + i * 3, 4, 3, 1].ColumnWidth = 18;
            }
            // Жирный шрифт
            xls[1, 3, 1 + stagesCount * 3, 3].Font.Bold = true;
            // Выравнивание
            xls[1, 3, 1 + stagesCount * 3, 3].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, 3, 1 + stagesCount * 3, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            // Границы вокруг ячеек
            xls[1, 3, 1 + stagesCount * 3, 3].BorderTable();
        }
        // Заголовок отчета + шапка таблицы (версия "в одну строку")
        private void CreateHeaderV1(int stagesCount, DateTime beginDate, DateTime endDate, ClassObject classObject, Xls xls)
        {
            // Заголовок отчета
            xls[2, 1].SetValue(String.Format("Состояния объектов РКК \"{0}\" за период с {1} по {2}", classObject, beginDate.ToShortDateString(), endDate.ToShortDateString()));
            xls[2, 1].Font.Bold = true;
            xls[2, 1].Font.Size = 14;

            // Шапка таблицы
            xls[1, 3].SetValue("Документ");
            xls[1, 3].ColumnWidth = 35;
            xls[1, 3, 1, 2].Merge();
            for (int i = 0; i < stagesCount; i++)
            {
                xls[2 + i * 5, 3].SetValue("Этап " + (i + 1));
                xls[2 + i * 5, 3, 5, 1].Merge();
                xls[2 + i * 5, 4].SetValue("Название");
                xls[3 + i * 5, 4].SetValue("Исполнители");
                xls[4 + i * 5, 4].SetValue("Время начала");
                xls[5 + i * 5, 4].SetValue("Время завершения");
                xls[6 + i * 5, 4].SetValue("Время этапа");
                xls[2 + i * 5, 4, 2, 1].ColumnWidth = 25;
                xls[4 + i * 5, 4, 5, 1].ColumnWidth = 18;
            }
            // Жирный шрифт
            xls[1, 3, 1 + stagesCount * 5, 2].Font.Bold = true;
            // Выравнивание
            xls[1, 3, 1 + stagesCount * 5, 2].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, 3, 1 + stagesCount * 5, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            // Границы вокруг ячеек
            xls[1, 3, 1 + stagesCount * 5, 2].BorderTable();
        }
        // Таблица отчета
        private void CreateTable(int startRow, Dictionary<string, List<ReportString>> reportTable, Xls xls)
        {
            int stagesCount = reportTable.Select(t => t.Value.Count).Max();
            int row = startRow;

            var progress = new ProgressDialog();
            progress.StartPosition = FormStartPosition.CenterScreen;
            progress.Show();

            progress.Text = "Формирование отчета";
            progress.SetInfo("Получение отчета Excel");
            var offset = 0;
            var count = reportTable.Count;
            foreach (var item in reportTable)
            {
                offset++;
                progress.SetInfo("Запись данных по документу: " + item.Key);
                progress.SetProgress(offset, count);

                xls[1, row].SetValue(item.Key);
                int cnt = 0;
                foreach (var value in item.Value)
                {
                    // Название этапа
                    xls[2 + cnt * 3, row].SetValue(value.Stage);
                    xls[2 + cnt * 3, row, 3, 1].Merge();
                    if (value.Stage.Length > 50)
                    {
                        xls[1, row].RowHeight = 30;
                        xls[2 + cnt * 3, row].WrapText = true;
                    }
                    // Исполнители
                    xls[2 + cnt * 3, row + 1].SetValue(String.Join(",", value.Executors.ToArray()));
                    xls[2 + cnt * 3, row + 1, 3, 1].Merge();
                    var stringCount = Math.Ceiling((double)String.Join(",", value.Executors.ToArray()).Length / 50);
                    if (String.Join(",", value.Executors.ToArray()).Length > 50)
                    {
                        xls[1, row + 1].RowHeight = 15 * stringCount;
                        xls[2 + cnt * 3, row + 1].WrapText = true;
                    }
                    // Установка времени начала, завершения этапа и интервала
                    xls[2 + cnt * 3, row + 2].SetValue(value.StartTime);
                    xls[3 + cnt * 3, row + 2].SetValue(value.EndTime);
                    xls[4 + cnt * 3, row + 2].SetValue(value.StageInterval);
                    // Выделение цветом, если конечный срок - пустой
                    if (String.IsNullOrEmpty(value.EndTime))
                        xls[2 + cnt * 3, row, 3, 3].Interior.Color = Color.LightGoldenrodYellow;
                    cnt++;
                }
                // Граница вокруг ячеек
                xls[1, row, 1, 3].BorderAround();
                for (int i = 0; i < stagesCount; i++) 
                {
                    xls[2 + i * 3, row, 3, 3].BorderAround();
                }
                // Выравнивание
                xls[1, row, 1 + stagesCount * 3, 3].VerticalAlignment = XlVAlign.xlVAlignTop;
                row += 3;
            }
            progress.Close();
        }
        // Таблица отчета (версия "в одну строку")
        private void CreateTableV1(int startRow, Dictionary<string, List<ReportString>> reportTable, Xls xls)
        {
            int stagesCount = reportTable.Select(t => t.Value.Count).Max();
            int row = startRow;

            foreach (var item in reportTable)
            {
                xls[1, row].SetValue(item.Key);
                int cnt = 0;
                foreach (var value in item.Value)
                {
                    // Название этапа
                    xls[2 + cnt * 5, row].SetValue(value.Stage);
                    // Исполнители
                    xls[3 + cnt * 5, row].SetValue(String.Join(",", value.Executors.ToArray()));
                    // Установка времени начала, завершения этапа и интервала
                    xls[4 + cnt * 5, row].SetValue(value.StartTime);
                    xls[5 + cnt * 5, row].SetValue(value.EndTime);
                    xls[6 + cnt * 5, row].SetValue(value.StageInterval);
                    // Выделение цветом, если конечный срок - пустой
                    if (String.IsNullOrEmpty(value.EndTime))
                        xls[2 + cnt * 5, row, 5, 1].Interior.Color = Color.LightGoldenrodYellow;
                    cnt++;
                }
                // Выравнивание
                xls[1, row, 1 + stagesCount * 5, 1].VerticalAlignment = XlVAlign.xlVAlignTop;
                row++;
            }
            // Границы вокруг ячеек
            xls[1, startRow, 1 + stagesCount * 5, row - startRow].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, startRow, 1, row - startRow].BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            for (int i = 0; i < stagesCount; i++)
            {
                xls[2 + i * 5, startRow, 5, row - startRow].BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }
            // Перенос текста
            xls[1, startRow, 1 + stagesCount * 5, row - startRow].WrapText = true;
        }
    }

    public class ReportString
    {
        public string Stage;                                       // Название этапа
        public List<string> Executors = new List<string>();        // Исполнители
        public string StartTime;                                   // Время начала этапа
        public string EndTime;                                     // Время завершения этапа
        public DateTime CurrentTime;                               // Текущее время (нужно для сортировки). Либо время завершения, либо текущее время
        public string StageInterval;                               // Время выполнения этапа
    }
}
