using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using ReportHelpers;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Calendar;
using TFlex.DOCs.Model.Search;
using TFlex.Reporting;
using Filter = TFlex.DOCs.Model.Search.Filter;

namespace WorkingTimeReport
{
    // отчет "Табель учета рабочего времени"
    public class WorkingTimeReport : IReportGenerator
    {

        public string DefaultReportFileExtension
        {
            get
            {
                return ".xls";
            }
        }

        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }


        public void Generate(IReportGenerationContext context)
        {
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

            var dialog = new ReportMainForm();
            if (dialog.ShowDialog(Form.ActiveForm) != DialogResult.OK)
            {
                return;
            }

            context.CopyTemplateFile();    // Создаем копию шаблона


            var department = dialog.departmentComboBox.SelectedItem as ReferenceObject;

            var workers = new List<ReferenceObject>();
            GetChilds(department, Guids.Cadres.Types.Worker, workers);

            workers = workers.OrderBy(t => t + "").ToList();

            Xls xls = new Xls();
            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);

                var date = dialog.date.Value;

                xls["V3"].SetValue(date.ToString("MMMM"));
                xls["Y3"].SetValue(date.ToString("yyyy") + " года");
                xls["K5"].SetValue(department + "");


                var timeBoardCadr = workers
                    .Where(t => t.Links.ToMany[Guids.Cadres.Links.AdditionalPosts].Any(t1 => t1[Guids.AdditionalPosts.Fields.Name].GetString() == "Табельщик"))
                    .FirstOrDefault();
                var chiefCadr = workers
                    .Where(t => t.Links.ToMany[Guids.Cadres.Links.AdditionalPosts].Any(t1 => t1[Guids.AdditionalPosts.Fields.Name].GetString() == "Начальник"))
                    .FirstOrDefault();
                xls["AC16"].SetValue(timeBoardCadr + "");
                xls["B16"].SetValue("Начальник " + department);
                xls["I16"].SetValue(chiefCadr + "");


                var rows = xls[1, 13].Rows.EntireRow;

                for (var i = 0; i < workers.Count - 1; i++)
                {
                    rows.Insert();
                    rows.Insert();
                }


                var calenarReference = new CalendarReference();
                var calendar = (CalendarReferenceObject)calenarReference.Objects[0];

                var unprovenTimeReferenceInfo = ReferenceCatalog.FindReference(Guids.UnprovenTime.RefereceGuid);
                var unprovenTimeReference = unprovenTimeReferenceInfo.CreateReference();
                var filter = new Filter(unprovenTimeReferenceInfo);

                var nextMonth = date.AddMonths(1);
                var firstDayOfNextMonth = new DateTime(nextMonth.Year, nextMonth.Month, 1, 0, 0, 0);
                var firstDayOfSelectedMonth = new DateTime(date.Year, date.Month, 1, 0, 0, 0);

                var term1 = filter.Terms.AddTerm("[" + Guids.UnprovenTime.Fields.StartDate + "]", ComparisonOperator.LessThan, firstDayOfNextMonth);
                term1.LogicalOperator = LogicalOperator.And;
                var term2 = filter.Terms.AddTerm("[" + Guids.UnprovenTime.Fields.EndDate + "]", ComparisonOperator.GreaterThan, firstDayOfSelectedMonth);
                term2.LogicalOperator = LogicalOperator.And;

                var unprovenTimes = unprovenTimeReference.Find(filter);

                var unprovenTimesByWorkers = new Dictionary<int, List<ReferenceObject>>();

                foreach (var unprovenTime in unprovenTimes)
                {
                    var worker = unprovenTime.GetObject(Guids.UnprovenTime.Links.Worker);
                    List<ReferenceObject> unptovenTimeList;
                    if (!unprovenTimesByWorkers.TryGetValue(worker.SystemFields.Id, out unptovenTimeList))
                    {
                        unptovenTimeList = new List<ReferenceObject>();
                        unprovenTimesByWorkers[worker.SystemFields.Id] = unptovenTimeList;
                    }
                    unptovenTimeList.Add(unprovenTime);
                }



                var days = DateTime.DaysInMonth(date.Year, date.Month);
                var daysInFirstLine = days / 2;
                var daysInSecondLine = days - daysInFirstLine;

                for (var day = 1; day <= days; day++)
                {
                    int col = 0;
                    int row = 0;

                    if (day <= daysInFirstLine)
                    {
                        col = 3 + day - 1;
                        row = 9;
                    }
                    else
                    {
                        col = 3 + day - 1 - daysInFirstLine;
                        row = 10;
                    }

                    xls[col, row].SetValue(day);

                    var dayDate = new DateTime(date.Year, date.Month, day);
                    var dayWorkTime = calendar.WorkTimeManager.GetRemainingTime(dayDate);

                    if (dayWorkTime.TotalMinutes == 0)
                    {
                        SetBackgroudGray(xls[col, row].Interior);
                        SetBackgroudGray(xls[col, row + 2].Interior);
                    }


                }


                xls[1, 11, 1, 2].EntireRow.Select();
                xls.Application.Selection.Copy();

                for (var i = 0; i < workers.Count - 1; i++)
                {
                    xls[1, 13 + i * 2, 1, 2].Rows.Select();
                    xls.Worksheet.Paste();
                }


                for (var i = 0; i < workers.Count; i++)
                {
                    var workerRow = 11 + i * 2;
                    xls[1, workerRow].SetValue(i + 1);
                    xls[2, workerRow].SetValue(workers[i] + "");

                    TimeSpan totalWorkerWorkTime = new TimeSpan();

                    int countWorkDays = 0;

                    List<ReferenceObject> unptovenTimeList;
                    if (!unprovenTimesByWorkers.TryGetValue(workers[i].SystemFields.Id, out unptovenTimeList))
                    {
                        unptovenTimeList = new List<ReferenceObject>();
                    }

                    int sickLeaveDaysCount = 0;

                    int[] leaveDaysCount = new int[15];


                    for (var day = 1; day <= days; day++)
                    {
                        int col = 0;
                        int row = 0;

                        if (day <= daysInFirstLine)
                        {
                            col = 3 + day - 1;
                            row = workerRow;
                        }
                        else
                        {
                            col = 3 + day - 1 - daysInFirstLine;
                            row = workerRow + 1;
                        }

                        var dayDate = new DateTime(date.Year, date.Month, day);
                        var dayWorkTime = calendar.WorkTimeManager.GetRemainingTime(dayDate);

                        ReferenceObject unprovenTimeOfThisDay = null;

                        foreach (var unprovenTime in unptovenTimeList)
                        {
                            var startDate = unprovenTime[Guids.UnprovenTime.Fields.StartDate].GetDateTime();
                            var endDate = unprovenTime[Guids.UnprovenTime.Fields.EndDate].GetDateTime();
                            startDate = new DateTime(startDate.Year, startDate.Month, startDate.Day, 0, 0, 0);

                            endDate = new DateTime(endDate.Year, endDate.Month, endDate.Day, 23, 59, 59);

                            var currentDate = new DateTime(dayDate.Year, dayDate.Month, dayDate.Day, 12, 0, 0);

                            if (currentDate <= endDate && currentDate >= startDate)
                            {
                                unprovenTimeOfThisDay = unprovenTime;
                                break;
                            }
                        }

                        if (unprovenTimeOfThisDay != null)
                        {
                            if (unprovenTimeOfThisDay.Class.IsInherit(Guids.UnprovenTime.Types.Leave))
                            {
                                var leaveType = unprovenTimeOfThisDay[Guids.UnprovenTime.LeaveFields.Type].GetInt32();
                                leaveDaysCount[leaveType]++;
                            }
                            else if (unprovenTimeOfThisDay.Class.IsInherit(Guids.UnprovenTime.Types.SickLeave))
                            {
                                sickLeaveDaysCount++;
                            }
                        }

                        if (dayWorkTime.TotalMinutes > 0)
                        {
                            // проврить больничные и отпуска и посчитать

                            if (unprovenTimeOfThisDay != null)
                            {
                                xls[col, row].SetValue("O");
                            }
                            else
                            {
                                xls[col, row].SetValue((int)dayWorkTime.TotalHours);
                                totalWorkerWorkTime += dayWorkTime;
                                countWorkDays++;
                            }
                        }
                        else
                        {
                            if (unprovenTimeOfThisDay != null)
                            {
                                xls[col, row].SetValue("O");
                            }
                        }
                    }
                    var rowStr = workerRow.ToString().Trim();
                    xls["S" + rowStr].SetValue(countWorkDays);

                    if (sickLeaveDaysCount > 0) xls["X" + rowStr].SetValue(sickLeaveDaysCount);

                    /*
                    Ежегодный отпуск	1	
                    Отпуск в связи с обучением	2	
                    Отпуск по беременности и родам	3	
                    Отпуск по уходу за ребенком до 1,5 лет	4	
                    Отпуск по уходу за ребенком до 3 лет	5	
                    Отпуск без сохранения заработной платы	6	
                    */
                    if (leaveDaysCount[1] > 0) xls["Z" + rowStr].SetValue(leaveDaysCount[1]);
                    if (leaveDaysCount[2] > 0) xls["W" + rowStr].SetValue(leaveDaysCount[2]);
                    if (leaveDaysCount[3] > 0) xls["Y" + rowStr].SetValue(leaveDaysCount[3]);
                    if (leaveDaysCount[4] > 0) xls["AA" + rowStr].SetValue(leaveDaysCount[4]);
                    if (leaveDaysCount[5] > 0) xls["AB" + rowStr].SetValue(leaveDaysCount[5]);
                    if (leaveDaysCount[6] > 0) xls["AC" + rowStr].SetValue(leaveDaysCount[6]);


                    xls["AJ" + rowStr].SetValue(workers[i][Guids.Cadres.WorkerFields.Post].GetString());
                    xls["AL" + rowStr].SetValue(workers[i][Guids.Cadres.WorkerFields.Pay].GetString());
                    xls["AM" + rowStr].SetValue(workers[i][Guids.Cadres.WorkerFields.TimeBoardNumber].GetString());
                    xls["AO" + rowStr].SetValue(Math.Round(totalWorkerWorkTime.TotalHours, 2));
                }

            }
            finally
            {
                xls.Quit(true);
                Process.Start(context.ReportFilePath);
            }
        }

        private static void SetBackgroudGray(Interior interior)
        {
            interior.Pattern = XlConstants.xlSolid;
            interior.PatternColorIndex = XlConstants.xlAutomatic;
            interior.ThemeColor = XlThemeColor.xlThemeColorAccent4;
            interior.TintAndShade = 0.799981688894314;
            interior.PatternTintAndShade = 0;
        }

        private void GetChilds(ReferenceObject parent, Guid guid, List<ReferenceObject> childs)
        {
            if (parent.Class.IsInherit(guid))
            {
                childs.Add(parent);
            }

            foreach (var child in parent.Children)
            {
                GetChilds(child, guid, childs);
            }
        }

    }
}
