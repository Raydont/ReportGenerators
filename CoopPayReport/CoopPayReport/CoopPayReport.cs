using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using ReportHelpers;
using TFlex.Reporting;
using Filter = TFlex.DOCs.Model.Search.Filter;
using System.Diagnostics;

namespace CoopPayReport
{
    // отчет "Учет затрат на кооперацию"
    public class CoopPayReport : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xlsx";
            }
        }

        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        public void Generate(IReportGenerationContext context)
        {
            // Cправочник "Учет договоров кооперации"
            Guid CoopAccountingGuid = new Guid("d0050b43-0277-4154-8b32-a60406b0328e");
            // Класс "Договор"
            Guid OrderClassGuid = new Guid("4fa11b9b-316f-48fa-a91b-e33d1950a813");
            // Класс "Оплата"
            Guid PayClassGuid = new Guid("971e612c-001f-4cdf-af3c-9990e9c471b3");
            // Класс "Поступление"
            Guid ArrivalClassGuid = new Guid("cc0b496e-944c-4fe4-a08e-e3c8493919cc");
            // Параметр "Дата документа"
            Guid DocDateGuid = new Guid("77c06eeb-b4e0-42b0-a117-5d51970a0e1f");
            // Параметр "Основание (документ)"
            Guid DocNameGuid = new Guid("157983ba-8a85-455f-be17-379b6cc948b9");
            // Параметр "Контрагент"
            Guid ContractorGuid = new Guid("5d40fb69-b7f6-4bfb-847d-4b2bc413fa8d");
            // Список объектов "Распределение по заказам"
            Guid OrderDistribObjectListGuid = new Guid("db7e790a-3be7-4a7f-846f-6c984ec225fa");
            // параметр "Заказ"
            Guid OrderNumGuid = new Guid("516f200b-25ab-4260-9030-b5f62d33a2f8");
            // параметр "Стоимость, руб"
            Guid OrderNumSumGuid = new Guid("5de26198-9a54-4dee-96d7-59d8c8ca8b9b");
            // параметр "Трудоемкость, н/ч"
            Guid OrderNumLabourGuid = new Guid("dcb7f278-a8ff-455f-8a20-154e82880d5a");


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

            var dialog = new mainPeriodForm();
            if (dialog.ShowDialog(Form.ActiveForm) != DialogResult.OK)
            {
                return;
            }

            context.CopyTemplateFile();    // Создаем копию шаблона

            ReferenceInfo coopAccountingInfo = ReferenceCatalog.FindReference(CoopAccountingGuid);
            Reference coopAccountingRef = coopAccountingInfo.CreateReference();
            ClassObject orderObjectClass = coopAccountingInfo.Classes.Find(OrderClassGuid);

            var orderDate = coopAccountingRef.ParameterGroup[DocDateGuid];

            DateTime startPeriod = dialog.startPeriod;
            DateTime endPeriod = dialog.endPeriod;

            Xls xls = new Xls();
            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);

                // наименование отчета
                xls["A1"].SetValue("Ведомость затрат по кооперации за " + GetPeriod(startPeriod, endPeriod));
                xls["A1"].Font.Bold = true;

                // заголовок таблицы отчета
                xls[1, 3].SetValue("Договор"); xls[1, 3, 1, 3].Merge();
                xls[2, 3].SetValue("Контрагент"); xls[2, 3, 1, 3].Merge();
                xls[3, 3].SetValue("Заказы"); xls[3, 3, 1, 3].Merge();
                xls[4, 3].SetValue("Сумма"); xls[4, 3, 2, 1].Merge();
                xls[4, 4].SetValue("руб."); xls[4, 4, 1, 2].Merge();
                xls[5, 4].SetValue("н/ч"); xls[5, 4, 1, 2].Merge();
                xls[6, 3].SetValue("Всего оплачено"); xls[6, 3, 2, 1].Merge();
                xls[6, 4].SetValue("руб."); xls[6, 4, 1, 2].Merge();
                xls[7, 4].SetValue("н/ч"); xls[7, 4, 1, 2].Merge();
                xls[8, 3].SetValue("Всего получено"); xls[8, 3, 2, 1].Merge();
                xls[8, 4].SetValue("руб."); xls[8, 4, 1, 2].Merge();
                xls[9, 4].SetValue("н/ч"); xls[9, 4, 1, 2].Merge();
                xls[10, 3].SetValue("Остаток по оплатам"); xls[10, 3, 2, 1].Merge();
                xls[10, 4].SetValue("руб."); xls[10, 4, 1, 2].Merge();
                xls[11, 4].SetValue("н/ч"); xls[11, 4, 1, 2].Merge();
                xls[12, 3].SetValue("Остаток по поступлениям"); xls[12, 3, 2, 1].Merge();
                xls[12, 4].SetValue("руб."); xls[12, 4, 1, 2].Merge();
                xls[13, 4].SetValue("н/ч"); xls[13, 4, 1, 2].Merge();

                // заголовки месяцев оплат и поступлений
                DateTime period = startPeriod;
                var monthCount = (endPeriod.Year - startPeriod.Year) * 12 + (endPeriod.Month - startPeriod.Month) + 1;
                int lastCol = 0;
                for (int i = 0; i < monthCount; i++)
                {
                    xls[14 + i * 5, 3].SetValue(period.ToString("MMMM")); xls[14 + i * 5, 3, 5, 1].Merge();
                    xls[14 + i * 5, 4].SetValue("Оплачено"); xls[14 + i * 5, 4, 3, 1].Merge();
                    xls[14 + i * 5, 5].SetValue("руб.");
                    xls[15 + i * 5, 5].SetValue("н/ч");
                    xls[16 + i * 5, 5].SetValue("Документ");
                    xls[17 + i * 5, 4].SetValue("Получение"); xls[17 + i * 5, 4, 2, 1].Merge();
                    xls[17 + i * 5, 5].SetValue("руб.");
                    xls[18 + i * 5, 5].SetValue("н/ч");

                    lastCol = 18 + i * 5;

                    period = period.AddMonths(1);
                }

                XlsExtender.BorderTable(xls[1, 3, lastCol, 3]);
                XlsExtender.CenterText(xls[1, 3, lastCol, 3]);

                // заполнение таблицы данными
                var filter = new Filter(coopAccountingInfo);
                filter.Terms.AddTerm(coopAccountingRef.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, orderObjectClass);
                filter.Terms.AddTerm(orderDate, ComparisonOperator.GreaterThanOrEqual, startPeriod);
                filter.Terms.AddTerm(orderDate, ComparisonOperator.LessThanOrEqual, endPeriod);
                var findedOrders = coopAccountingRef.Find(filter);

                int row = 6;
                int lastRow = 0;
                foreach (var order in findedOrders)
                {
                    // договор
                    xls[1, row].SetValue(order[DocNameGuid].GetString() + " от " + order[DocDateGuid].GetDateTime().ToShortDateString());
                    xls[2, row].SetValue(order[ContractorGuid].GetString());
                    int orderRow = row;
                    var ordersDistrib = order.GetObjects(OrderDistribObjectListGuid).OrderBy(t => t[OrderNumGuid].GetString());
                    // заказы договора и распределение сумм и трудоемкости по ним
                    foreach (var orderDistribItem in ordersDistrib)
                    {
                        xls[3, row].SetValue("'" + orderDistribItem[OrderNumGuid].GetString());
                        xls[4, row].SetValue(orderDistribItem[OrderNumSumGuid].GetDecimal());
                        xls[5, row].SetValue(orderDistribItem[OrderNumLabourGuid].GetDouble());
                        row++;
                    }
                    // оплаты и поступления по договору
                    var payArrivalCollection = order.Children;
                    foreach (var payArrival in payArrivalCollection)
                    {
                        var payArrivalDate = payArrival[DocDateGuid].GetDateTime();
                        var monthPosition = payArrivalDate.Month + (payArrivalDate.Year - startPeriod.Year) * 12 - (startPeriod.Month - 1);
                        int distribRow = orderRow;
                        var payArrivalsDistrib = payArrival.GetObjects(OrderDistribObjectListGuid).OrderBy(t => t[OrderNumGuid].GetString());
                        foreach (var orderDistrib in payArrivalsDistrib)
                        {
                            if (payArrival.Class.Guid == PayClassGuid)
                            {
                                xls[14 + (monthPosition - 1) * 5, distribRow].SetValue(orderDistrib[OrderNumSumGuid].GetDecimal());
                                xls[15 + (monthPosition - 1) * 5, distribRow].SetValue(orderDistrib[OrderNumLabourGuid].GetDouble());
                                xls[16 + (monthPosition - 1) * 5, orderRow].SetValue(payArrival[DocNameGuid].Value + " от " + payArrival[DocDateGuid].GetDateTime().ToShortDateString());
                                distribRow++;
                            }
                            else
                            {
                                xls[17 + (monthPosition - 1) * 5, distribRow].SetValue(orderDistrib[OrderNumSumGuid].GetDecimal());
                                xls[18 + (monthPosition - 1) * 5, distribRow].SetValue(orderDistrib[OrderNumLabourGuid].GetDouble());
                                distribRow++;
                            }
                        }
                    }

                    if (order.GetObjects(OrderDistribObjectListGuid).Count == 0) row++;

                    // граница разделения договоров в таблице
                    xls[1, row, lastCol, 1].Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;

                    lastRow = row;
                }

                // формулы:

                // сумма оплат, поступлений
                List<string> paySumCells = new List<string>();
                List<string> payLabourCells = new List<string>();
                List<string> arrivalSumCells = new List<string>();
                List<string> arrivalLabourCells = new List<string>();
                for (int i = 0; i < monthCount; i++)
                {
                    paySumCells.Add(Xls.Letter(14 + 5 * i) + "6");
                    payLabourCells.Add(Xls.Letter(15 + 5 * i) + "6");
                    arrivalSumCells.Add(Xls.Letter(17 + 5 * i) + "6");
                    arrivalLabourCells.Add(Xls.Letter(18 + 5 * i) + "6");
                }
                XlsExtender.SetFormula(xls[6, 6], "=" + String.Join("+", paySumCells));
                XlsExtender.SetFormula(xls[7, 6], "=" + String.Join("+", payLabourCells));
                XlsExtender.SetFormula(xls[8, 6], "=" + String.Join("+", arrivalSumCells));
                XlsExtender.SetFormula(xls[9, 6], "=" + String.Join("+", arrivalLabourCells));
                xls[6, 6].AutoFill(xls[6, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);    // "размножение" формул
                xls[7, 6].AutoFill(xls[7, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);
                xls[8, 6].AutoFill(xls[8, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);
                xls[9, 6].AutoFill(xls[9, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);

                // суммы остатков оплат и поступлений
                XlsExtender.SetFormula(xls[10, 6], "=D6-F6"); // остаток оплат (сумма)
                XlsExtender.SetFormula(xls[11, 6], "=E6-G6"); // остаток оплат (трудоемкость)
                XlsExtender.SetFormula(xls[12, 6], "=D6-H6"); // остаток поступлений (сумма)
                XlsExtender.SetFormula(xls[13, 6], "=E6-I6"); // остаток поступлений (трудоемкость)
                xls[10, 6].AutoFill(xls[10, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);    // "размножение" формул
                xls[11, 6].AutoFill(xls[11, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);
                xls[12, 6].AutoFill(xls[12, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);
                xls[13, 6].AutoFill(xls[13, 6, 1, lastRow - 6], XlAutoFillType.xlFillValues);
            }
            finally
            {
                xls.Quit(true);
                Process.Start(context.ReportFilePath);
            }

        }

        private string GetPeriod(DateTime startPeriod, DateTime endPeriod)
        {
            if (startPeriod.Day == 1 && startPeriod.Month == 1
                && endPeriod.Day == 31 && endPeriod.Month == 12)
            {
                return startPeriod.Year + " год";
            }
            else
            {
                return "период c " + startPeriod.ToString("dd MMMM yyyy") + " по " + endPeriod.ToString("dd MMMM yyyy");
            }
        }
    }
}
