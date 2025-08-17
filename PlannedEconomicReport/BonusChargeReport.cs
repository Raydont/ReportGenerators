using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ReportHelpers;
using Slepov.Russian.СуммаПрописью;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Calendar;
using TFlex.DOCs.Model.Search;
using TFlex.Reporting;
using Filter = TFlex.DOCs.Model.Search.Filter;

namespace PlannedEconomicReport
{
    // отчет "Ведомость начисления премии работникам"
    class BonusChargeReport : IReportGenerator
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

            var dialog = new BonusChargeMainForm();
            if (dialog.ShowDialog(Form.ActiveForm) != DialogResult.OK)
            {
                return;
            }

            context.CopyTemplateFile();    // Создаем копию шаблона

            var department = dialog.departmentComboBox.SelectedItem as ReferenceObject;
            var workers = new List<ReferenceObject>();
            CommonClasses.GetChilds(department, Guids.Cadres.Types.Worker, workers);
            
            var chiefCadr = workers
                      .Where(t => t.Links.ToMany[Guids.Cadres.Links.AdditionalPosts].Any(t1 => t1[Guids.AdditionalPosts.Fields.Name].GetString() == "Начальник"))
                      .FirstOrDefault();

            workers = workers.OrderBy(t => t + "").Where(t => t != chiefCadr).ToList();

            Xls xls = new Xls();
            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);

                xls["G1"].SetValue(department.ToString());

                var date = dialog.date.Value;
                xls["A11"].SetValue("за " + date.ToString("MMMM") + " " + date.ToString("yyyy") + " года");

                var orderNumber = dialog.orderNumber.Text;
                var orderDate = dialog.orderDate.Value;
                xls["A12"].SetValue("В соответствии с приказом № " + orderNumber + " от " + orderDate.ToString("dd MMMM yyyy"));

                xls["A23"].SetValue("Начальник " + department.ToString().ToLower() + "____________________ /" + chiefCadr + "/");

                var rows = xls[1, 18].Rows.EntireRow;

                for (var i = 0; i < workers.Count - 1; i++)
                {
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
                var totalWorkingDays = 0;
                for (var day = 1; day <= days; day++)
                {
                    var dayDate = new DateTime(date.Year, date.Month, day);
                    var dayWorkTime = calendar.WorkTimeManager.GetRemainingTime(dayDate);

                    if (dayWorkTime.TotalMinutes > 0)
                    {
                        totalWorkingDays++;
                    }
                }


                for (var i = 0; i < workers.Count; i++)
                {
                    var workerRow = 17 + i;
                    xls[1, workerRow].SetValue(i + 1);
                    xls[2, workerRow].SetValue(workers[i] + "");
                    xls[3, workerRow].SetValue(workers[i][Guids.Cadres.WorkerFields.Post].GetString());
                    xls[4, workerRow].SetValue(workers[i][Guids.Cadres.WorkerFields.Pay].GetString());
                    var bonusPercent = Convert.ToInt32(dialog.bonusPercent.Value);
                    xls[7, workerRow].SetValue(bonusPercent.ToString());

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
                            // проверить больничные и отпуска и посчитать

                            if (unprovenTimeOfThisDay == null)
                            {
                                totalWorkerWorkTime += dayWorkTime;
                                countWorkDays++;
                            }
                        }
                    }
                    xls[5, workerRow].SetValue(countWorkDays);
                    xls[6, workerRow].SetFormula("=D" + workerRow + "/" + totalWorkingDays + "*" + "E" + workerRow);
                    xls[8, workerRow].SetFormula("=ОКРУГЛВВЕРХ((F" + workerRow + "*G" + workerRow + "/100)/5;0)*5");
                    xls[1, workerRow].EntireRow.AutoFit();
                }
                var totalSumRow = 20 + workers.Count - 1;
                var bonusSum = Convert.ToDecimal(xls[8, totalSumRow].Value);
                var sumInWords = new StringBuilder(Валюта.Рубли.Пропись(bonusSum));
                Заглавные.Первая.Применить(sumInWords);
                string sumInWordsString = sumInWords.ToString().Replace(" 00 копеек", ".").Trim();
                xls["E3"].SetValue("Сумма: " + sumInWordsString);
            }
            finally
            {
                xls.Quit(true);
                Process.Start(context.ReportFilePath);
            }
        }
    }
}
