using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportHelpers;
using System.Globalization;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Text.RegularExpressions;
using System.Drawing;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.Macros;
using System.Xml.Linq;

namespace DocumentsListReport
{
    public class ReportTables
    {
        MacroProvider macro;
        public ReportTables(MacroProvider macro)
        {
            this.macro = macro;
        }
        public void OrderDocTable(Xls xls, int row, Объекты objects)
        {
            var excelOps = new ExcelOperations();
            int rownach = row;
            macro.ДиалогОжидания.Показать("Идет формирование отчета...", true);
            // Упорядочивание по дате регистрации и номеру документа
            Regex numReg = new Regex(@"\d+");
            objects = objects.OrderBy(t => ((DateTime)t[Guids.RegCardsReference.Fields.regDate.ToString()])).
                ThenBy(t => Convert.ToInt32(numReg.Match(t[Guids.RegCardsReference.Fields.regNumber.ToString()].ToString()).Value)).ToList();
            excelOps.PageSetup(xls);
            int cnt = 0;
            foreach (var item in objects)
            {
                // Номер приказа
                xls[1, row].NumberFormat = "@";
                xls[1, row].WrapText = true;

                String docNumber = item[Guids.RegCardsReference.Fields.regNumber.ToString()];
                if (docNumber.StartsWith("Приказ "))
                    docNumber = docNumber.Replace("Приказ ", "");
                String docData = ((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToShortDateString();
                xls[1, row].SetValue(docNumber);
                // Дата регистрации
                xls[2, row].SetValue(docData);
                // Краткое содержание
                xls[3, row].SetValue(item[Guids.RegCardsReference.OrderDoc.Fields.content.ToString()].ToString());
                // Подписал
                xls[4, row].SetValue(item[Guids.RegCardsReference.OrderDoc.Fields.signed.ToString()].ToString());
                // Исполнитель
                if (item[Guids.RegCardsReference.Links.executorLink.ToString()] != null)
                    xls[5, row].SetValue(item[Guids.RegCardsReference.Links.executorLink.ToString()].ToString());
                xls[1, row, 5, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
                row +=2;
                cnt++;
                double percent = objects.Count == 0 ? 0 : (double)cnt / objects.Count * 100;
                macro.ДиалогОжидания.СледующийШаг("Идет запись документа " + docNumber + " от " + docData, percent);
            }
            xls[1, rownach, 5, row - rownach].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, rownach, 1, row - rownach].HorizontalAlignment = XlHAlign.xlHAlignRight;
            xls[3, rownach, 3, row - rownach].WrapText = true;
            for (int i = 1; i <= 5; i++)
                xls[i, rownach, 1, row - rownach].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft,
                                                                                                              XlBordersIndex.xlEdgeRight);
            macro.ДиалогОжидания.Скрыть();
        }
        public void DecreeDocTable(Xls xls, int row, Объекты objects)
        {
            int rownach = row;
            macro.ДиалогОжидания.Показать("Идет формирование отчета...", true);
            // Упорядочивание по дате регистрации и номеру документа
            Regex numReg = new Regex(@"\d+");
            objects = objects.OrderBy(t => ((DateTime)t[Guids.RegCardsReference.Fields.regDate.ToString()])).
                ThenBy(t => Convert.ToInt32(numReg.Match(t[Guids.RegCardsReference.Fields.regNumber.ToString()].ToString()).Value)).ToList();
            int cnt = 0;
            foreach (var item in objects)
            {
                // Номер приказа
                xls[1, row].NumberFormat = "@";
                xls[1, row].WrapText = true;

                String docNumber = item[Guids.RegCardsReference.Fields.regNumber.ToString()];
                if (docNumber.StartsWith("Распоряжение "))
                    docNumber = docNumber.Replace("Распоряжение ", "");
                String docData = ((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToShortDateString();

                xls[1, row].SetValue(docNumber);
                // Дата регистрации
                xls[2, row].SetValue(docData);
                // Краткое содержание
                xls[3, row].SetValue(item[Guids.RegCardsReference.OrderDoc.Fields.content.ToString()].ToString());
                // Подписал
                xls[4, row].SetValue(item[Guids.RegCardsReference.OrderDoc.Fields.signed.ToString()].ToString());
                // Исполнитель
                if (item[Guids.RegCardsReference.Links.executorLink.ToString()] != null)
                    xls[5, row].SetValue(item[Guids.RegCardsReference.Links.executorLink.ToString()].ToString());
                xls[1, row, 5, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
                // Последняя строчка
                row +=2;
                cnt++;
                double percent = objects.Count == 0 ? 0 : (double)cnt / objects.Count * 100;
                // Задержка по времени при маленькой коллекции
                if (objects.Count < 5)
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                macro.ДиалогОжидания.СледующийШаг("Идет запись документа " + docNumber + " от " + docData, percent);
            }
            xls[1, rownach, 5, row - rownach].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, rownach, 1, row - rownach].HorizontalAlignment = XlHAlign.xlHAlignRight;
            xls[3, rownach, 3, row - rownach].WrapText = true;
            for (int i = 1; i <= 5; i++)
                xls[i, rownach, 1, row - rownach].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft,
                                                                                                              XlBordersIndex.xlEdgeRight);
            macro.ДиалогОжидания.Скрыть();
        }
        public void InputDocTable(Xls xls, int row, Объекты objects)
        {
            var excelOps = new ExcelOperations();
            int rowmax = row;
            int rownach = row;
            int row1;
            macro.ДиалогОжидания.Показать("Идет формирование отчета...", true);
            // Упорядочивание номеру документа и дате регистрации
            Regex numReg = new Regex(@"\d+");
            objects = objects.OrderBy(t => Convert.ToInt32(numReg.Match(t[Guids.RegCardsReference.InputDoc.Fields.orderNumber.ToString()].ToString()).Value)).
                ThenBy(t => ((DateTime)t[Guids.RegCardsReference.Fields.regDate.ToString()])).ToList();
            excelOps.PageSetup(xls);
            int cnt = 0;
            foreach (var item in objects)
            {
                // Порядковый номер
                xls[1, row].NumberFormat = "@";
                xls[1, row].WrapText = true;
                xls[1, row].SetValue(item[Guids.RegCardsReference.InputDoc.Fields.orderNumber.ToString()]);
                // Список адресатов
                row1 = row;
                foreach (var receiver in item.СвязанныеОбъекты[Guids.RegCardsReference.InputDoc.Links.receiversLink.ToString()])
                {
                    xls[2, row1].SetValue(receiver[Guids.RegCardsReference.InputDoc.Receivers.receicerName.ToString()] + " (" +
                        receiver[Guids.RegCardsReference.InputDoc.Receivers.receiverCity.ToString()] + ")");
                    row1++;
                }
                if (row1 > rowmax) rowmax = row1;
                // Номер исходящего
                xls[3, row].SetValue(item[Guids.RegCardsReference.Fields.regNumber.ToString()]);
                // Краткое содержание
                xls[4, row].SetValue(item[Guids.RegCardsReference.InputDoc.Fields.summary.ToString()]);
                xls[1, row, 4, rowmax - row + 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
                // Последняя строчка блока текущего документа
                if (rowmax > row)
                    row = rowmax;
                row++;
                String docData = ((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToShortDateString();
                String docNumber = item[Guids.RegCardsReference.Fields.regNumber.ToString()];
                cnt++;
                double percent = objects.Count == 0 ? 0 : (double)cnt / objects.Count * 100;
                if (objects.Count < 5)
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                macro.ДиалогОжидания.СледующийШаг("Идет запись документа " + docNumber + " от " + docData, percent);
            }
            xls[1, rownach, 4, row - rownach].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[2, rownach, 1, row - rownach].WrapText = true;
            xls[4, rownach, 1, row - rownach].WrapText = true;
            for (int i = 1; i <= 4; i++)
                xls[i, rownach, 1, row - rownach].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft,
                                                                                                              XlBordersIndex.xlEdgeRight);
            macro.ДиалогОжидания.Скрыть();
        }
        public void ContractDocTable(Xls xls, int row, Объекты objects)
        {
            var excelOps = new ExcelOperations();
            int rownach = row;
            int row1;
            macro.ДиалогОжидания.Показать("Идет формирование отчета...", true);
            excelOps.PageSetup(xls);
            // Упорядочивание по дате регистрации и номеру документа
            Regex numReg = new Regex(@"\d+");
            objects = objects.OrderBy(t => ((DateTime)t[Guids.RegCardsReference.Fields.regDate.ToString()])).
                ThenBy(t => Convert.ToInt32(numReg.Match(t[Guids.RegCardsReference.Fields.regNumber.ToString()].ToString()).Value)).ToList();
            int cnt = 0;
            foreach (var item in objects)
            {
                xls[3, row, 3, 1].WrapText = true;
                // От
                String docData = ((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToShortDateString();
                xls[1, row].SetValue(((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToString("dd.MM.yyyy"));
                // Номер ДЗУ
                String docNumber = item[Guids.RegCardsReference.Fields.regNumber.ToString()];
                xls[2, row].SetValue(docNumber);
                // Краткое содержание
                xls[3, row].SetValue(item[Guids.RegCardsReference.ContractDoc.Fields.content.ToString().ToString()]);
                // Номер договора
                xls[4, row].SetValue(item[Guids.RegCardsReference.ContractDoc.Fields.contractNum.ToString().ToString()]);
                // Контрагент
                xls[5, row].SetValue(item[Guids.RegCardsReference.ContractDoc.Fields.contractor.ToString().ToString()]);

                xls[1, row, 6, 1].Font.Bold = true;
                xls[1, row, 6, 1].Font.Color = Color.Blue;

                row1 = row;
                // Поиск объектов, у которых в свойстве "Рамочный договор" стоит текущий объект
                string условие = String.Format("[Тип] = '{0}' И [{1}]->[Guid] = '{2}'", Guids.RegCardsReference.ContractDoc.type.ToString(), 
                    Guids.RegCardsReference.ContractDoc.Links.frameContractLink.ToString(), ((ReferenceObject)item).SystemFields.Guid);
                var linkedObjects = macro.НайтиОбъекты(Guids.RegCardsReference.refGuid.ToString(), условие);
                foreach (var linkedDZU in linkedObjects)
                {
                    row1++;
                    xls[3, row1, 3, 1].WrapText = true;
                    // От
                    xls[1, row1].SetValue(((DateTime)linkedDZU[Guids.RegCardsReference.Fields.regDate.ToString()]).ToString("dd.MM.yyyy"));
                    // Номер ДЗУ
                    xls[2, row1].SetValue(linkedDZU[Guids.RegCardsReference.Fields.regNumber.ToString()]);
                    // Краткое содержание
                    xls[3, row1].SetValue(linkedDZU[Guids.RegCardsReference.ContractDoc.Fields.content.ToString()]);
                    // Номер договора
                    xls[4, row1].SetValue(linkedDZU[Guids.RegCardsReference.ContractDoc.Fields.contractNum.ToString()]);
                    // Контрагент
                    xls[5, row1].SetValue(linkedDZU[Guids.RegCardsReference.ContractDoc.Fields.contractor.ToString()]);
                    // Сумма с НДС
                    xls[6, row1].SetValue(linkedDZU[Guids.RegCardsReference.ContractDoc.Fields.sumWithNDS.ToString()]);
                }
                xls[1, row, 6, row1 + 2 - row].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
                row = row1 + 2;
                cnt++;
                double percent = objects.Count == 0 ? 0 : (double)cnt / objects.Count * 100;
                if (objects.Count < 5)
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                macro.ДиалогОжидания.СледующийШаг("Идет запись документа " + docNumber + " от " + docData, percent);
            }
            xls[1, 1, 6, row + 1].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[4, rownach, 1, row + 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            for (int i = 1; i <= 6; i++)
                xls[i, rownach, 1, row - rownach].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft,
                                                                                                              XlBordersIndex.xlEdgeRight);
            macro.ДиалогОжидания.Скрыть();
        }
        public void ServiceDocTable(Xls xls, int row, Объекты objects)
        {
            var excelOps = new ExcelOperations();
            int rowmax = row;
            int rownach = row;
            int row1;
            macro.ДиалогОжидания.Показать("Идет формирование отчета...", true);
            excelOps.PageSetup(xls);
            // Упорядочивание по дате регистрации и номеру документа
            Regex numReg = new Regex(@"\d+");
            objects = objects.OrderBy(t => ((DateTime)t[Guids.RegCardsReference.Fields.regDate.ToString()])).
                ThenBy(t => Convert.ToInt32(numReg.Match(t[Guids.RegCardsReference.Fields.regNumber.ToString()].ToString()).Value)).ToList();
            int cnt = 0;
            foreach (var item in objects)
            {
                row++;
                // Регистрационный номер СЗ
                xls[1, row].NumberFormat = "@";
                String docNumber = item[Guids.RegCardsReference.Fields.regNumber.ToString()];
                xls[1, row].SetValue(docNumber.ToString().Replace("СЗ ", ""));
                // Дата регистрации
                String docData = ((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToShortDateString();
                xls[2, row].SetValue(docData);
                // Исполнитель
                xls[3, row].SetValue(item.Автор.ToString());
                // Краткое содержание
                xls[4, row].WrapText = true;
                if (item[Guids.RegCardsReference.Fields.content.ToString()] != null)
                    xls[4, row].SetValue(item[Guids.RegCardsReference.Fields.content.ToString()]);
                else
                    xls[4, row].SetValue(String.Empty);
                // Список получателей
                row1 = row;
                if (item.СвязанныеОбъекты[Guids.RegCardsReference.Links.receiversLink.ToString()] != null)
                {
                    foreach (var receiver in item.СвязанныеОбъекты[Guids.RegCardsReference.Links.receiversLink.ToString()])
                    {
                        xls[5, row1].SetValue(receiver.ToString());
                        row1++;
                    }
                }
                else
                    xls[5, row].SetValue(String.Empty);

                if (row1 > rowmax) rowmax = row1;

                // Список визирующих лиц
                row1 = row;
                if (item.СвязанныеОбъекты[Guids.RegCardsReference.Links.rCUsersLink.ToString()] != null)
                {
                    foreach (var visor in item.СвязанныеОбъекты[Guids.RegCardsReference.Links.rCUsersLink.ToString()])
                    {
                        xls[6, row1].SetValue(visor.ToString());
                        row1++;
                    }
                }
                else
                    xls[6, row].SetValue(String.Empty);

                if (row1 > rowmax) rowmax = row1;

                // Стадия документа
                xls[7, row].SetValue(((ReferenceObject)item).SystemFields.Stage.ToString());

                xls[1, row, 7, rowmax - row + 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
                // Последняя строчка
                if (rowmax > row)
                    row = rowmax;
                else
                    row++;
                cnt++;
                double percent = objects.Count == 0 ? 0 : (double)cnt / objects.Count * 100;
                if (objects.Count < 5)
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                macro.ДиалогОжидания.СледующийШаг("Идет запись документа " + docNumber + " от " + docData, percent);
            }

            xls[1, rownach, 7, row - rownach].VerticalAlignment = XlVAlign.xlVAlignTop;
            for (int i = 1; i <= 7; i++)
                xls[i, rownach - 1, 1, row - rownach + 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft,
                                                                                                                      XlBordersIndex.xlEdgeRight);
            macro.ДиалогОжидания.Скрыть();
        }
        public void ReportDocTable(Xls xls, int row, Объекты objects)
        {
            var excelOps = new ExcelOperations();
            int rowmax = row;
            int rownach = row;
            int row1;
            macro.ДиалогОжидания.Показать("Идет формирование отчета...", true);
            // Упорядочивание по дате регистрации и номеру документа
            Regex numReg = new Regex(@"\d+");
            objects = objects.OrderBy(t => ((DateTime)t[Guids.RegCardsReference.Fields.regDate.ToString()])).
                ThenBy(t => Convert.ToInt32(numReg.Match(t[Guids.RegCardsReference.Fields.regNumber.ToString()].ToString()).Value)).ToList();
            excelOps.PageSetup(xls);
            int cnt = 0;
            foreach (var item in objects)
            {
                row++;
                // Регистрационный номер ДЗ
                xls[1, row].NumberFormat = "@";
                String docNumber = item[Guids.RegCardsReference.Fields.regNumber.ToString()];
                xls[1, row].SetValue(docNumber.ToString().Replace("ДЗ ", ""));
                // Дата регистрации
                String docData = ((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToShortDateString();
                xls[2, row].SetValue(docData);
                // Исполнитель
                xls[3, row].SetValue(item.Author.ToString());
                // Краткое содержание
                xls[4, row].WrapText = true;
                if (item[Guids.RegCardsReference.Fields.content.ToString()] != null)
                    xls[4, row].SetValue(item[Guids.RegCardsReference.Fields.content.ToString()]);
                else
                    xls[4, row].SetValue(String.Empty);
                // Список получателей
                row1 = row;
                if (item.СвязанныеОбъекты[Guids.RegCardsReference.Links.receiversLink.ToString()] != null)
                {
                    foreach (var receiver in item.СвязанныеОбъекты[Guids.RegCardsReference.Links.receiversLink.ToString()])
                    {
                        xls[5, row1].WrapText = true;
                        xls[5, row1].SetValue(receiver.ToString());
                        row1++;
                    }
                }
                else
                {
                    xls[5, row].WrapText = true;
                    xls[5, row].SetValue(String.Empty);
                }

                if (row1 > rowmax) rowmax = row1;

                // Подписал
                xls[6, row].SetValue(item[Guids.RegCardsReference.ReportDoc.Fields.sender.ToString()]);

                // Подразделение
                xls[7, row].NumberFormat = "@";
                xls[7, row].SetValue(item[Guids.RegCardsReference.Fields.departmentNumber.ToString()]);

                // Стадия документа
                xls[8, row].SetValue(((ReferenceObject)item).SystemFields.Stage.ToString());

                // Последняя строчка
                if (rowmax > row)
                    row = rowmax;
                else
                    row++;

                xls[1, row, 8, rowmax - row + 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
                cnt++;
                double percent = objects.Count == 0 ? 0 : (double)cnt / objects.Count * 100;
                if (objects.Count < 5)
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                macro.ДиалогОжидания.СледующийШаг("Идет запись документа " + docNumber + " от " + docData, percent);
            }

            xls[1, rownach, 8, row - rownach].VerticalAlignment = XlVAlign.xlVAlignTop;
            for (int i = 1; i <= 8; i++)
                xls[i, rownach - 1, 1, row - rownach + 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft,
                                                                                                                      XlBordersIndex.xlEdgeRight);
            macro.ДиалогОжидания.Скрыть();
        }
        public void PaymentServiceDocTable(Xls xls, int row, Объекты objects, bool isPaid)
        {
            var excelOps = new ExcelOperations();
            int rownach = row;
            int rowmax = row;
            int row1;
            // Упорядочивание по дате регистрации и номеру документа
            Regex numReg = new Regex(@"\d+");
            objects = objects.OrderBy(t => ((DateTime)t[Guids.RegCardsReference.Fields.regDate.ToString()])).
                ThenBy(t => Convert.ToInt32(numReg.Match(t[Guids.RegCardsReference.Fields.regNumber.ToString()].ToString()).Value)).ToList();
            excelOps.PageSetup(xls);
            int cnt = 0;
            macro.ДиалогОжидания.Показать("Идет формирование отчета...", true);
            foreach (var item in objects)
            {
                row++;
                cnt++;
                // Номер пункта
                xls[1, row].SetValue(cnt);
                // Регистрационный номер СЗНО
                xls[2, row].NumberFormat = "@";
                String docNumber = item[Guids.RegCardsReference.Fields.regNumber.ToString()];
                xls[2, row].SetValue(docNumber.Replace("СЗНО/ЗПЗП ", ""));
                // Дата регистрации
                String docData = ((DateTime)item[Guids.RegCardsReference.Fields.regDate.ToString()]).ToShortDateString();
                xls[3, row].SetValue(docData);
                // Контрагент
                xls[4, row].WrapText = true;
                xls[4, row].SetValue(item[Guids.RegCardsReference.PaymentServiceDoc.Fields.contractor.ToString()]);
                // Краткое содержание
                xls[5, row].WrapText = true;
                xls[5, row].SetValue(item[Guids.RegCardsReference.Fields.content.ToString()]);
                // Статьи БДДС
                row1 = row;
                var articles = item.СвязанныеОбъекты[Guids.RegCardsReference.PaymentServiceDoc.Links.budgetArticlesListLink.ToString()];
                if (articles != null && articles.Count > 0)
                {
                    // Если статьи БДДС заполнены в таблице, то сумма к оплате берется из таблицы
                    foreach (var article in articles)
                    {
                        xls[6, row1].SetValue(article[Guids.RegCardsReference.PaymentServiceDoc.BudgetArticlesList.budgetArticleCode.ToString()]);
                        xls[7, row1].NumberFormat = "#,##0.00 $";
                        xls[7, row1].SetValue(article[Guids.RegCardsReference.PaymentServiceDoc.BudgetArticlesList.paymentSum.ToString()]);
                        row1++;
                    }
                }
                else
                {
                    // Иначе сумма к оплате берется из поля "Сумма к оплате"
                    xls[7, row].NumberFormat = "#,##0.00 $";
                    xls[7, row].SetValue(item[Guids.RegCardsReference.PaymentServiceDoc.Fields.paymentSum.ToString()]);
                }
                if (row1 > rowmax) rowmax = row1;
                // Номер договора
                string contractNumber = item[String.Format("[{0}]->[{1}]", Guids.RegCardsReference.PaymentServiceDoc.Links.contractLink.ToString(),
                                                                           Guids.RegCardsReference.Fields.regNumber.ToString())];
                xls[8, row].WrapText = true;
                xls[8, row].SetValue(contractNumber.Replace("ДЗУ ", ""));
                // Номер заказа
                xls[9, row].WrapText = true;
                xls[9, row].SetValue(item[String.Format("[{0}]->[{1}]", Guids.RegCardsReference.PaymentServiceDoc.Links.contractLink.ToString(),
                                                                        Guids.RegCardsReference.ContractDoc.Fields.orderNum.ToString())]);
                if (isPaid)
                    xls[10, row].SetValue("Оплачено");
                else
                    xls[10, row].SetValue(item["Стадия"]);

                // Последняя строчка
                if (rowmax > row)
                    row = rowmax;
                else
                    row++;

                xls[1, row, 10, rowmax - row + 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
                double percent = objects.Count == 0 ? 0 : (double)cnt / objects.Count * 100;
                if (objects.Count < 5)
                    System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                macro.ДиалогОжидания.СледующийШаг("Идет запись документа " + docNumber + " от " + docData, percent);

            }
            xls[1, rownach, 10, row - rownach].VerticalAlignment = XlVAlign.xlVAlignTop;
            for (int i = 1; i <= 10; i++)
                xls[i, rownach - 1, 1, row - rownach + 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft,
                                                                                                              XlBordersIndex.xlEdgeRight);
            macro.ДиалогОжидания.Скрыть();
            // Заполнение итоговых сумм по статьям БДДС
            int rowCounter = ArticlesSumm(objects, row, xls);

            var cell = Xls.Address(7, rownach + 1);
            string endCell = Xls.Address(7, row);

            // Вычисление итоговой сумма к оплате (с учетом количества строк под суммы по статьям бюджета)
            string value = "=СУММ(" + cell + ":" + endCell + ")";
            int sumRow = row + rowCounter + 1;
            xls[7, sumRow].NumberFormat = "#,##0.00 $";
            xls[7, sumRow].SetFormula(value);

            // Запись "Итого платежей:"
            xls[5, sumRow].SetValue("Всего платежей:");

            // Границы ячеек кода статьи и суммы
            xls[6, sumRow, 2, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                                 XlBordersIndex.xlInsideVertical);
            // Рамки ячеек до и после кода статьи и суммы
            xls[1, sumRow, 5, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                                 XlBordersIndex.xlInsideHorizontal);

            xls[8, sumRow, 3, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                                 XlBordersIndex.xlInsideHorizontal);
        }
        // Заполнение сумм по кодам статей БДДС (для СЗНО)
        private int ArticlesSumm(Объекты objects, int row, Xls xls)
        {
            int summaryRow = 0;
            int rowCounter = 0;

            // Список кодов статей бюджета
            var articleCodes = objects.Where(t => t.СвязанныеОбъекты[Guids.RegCardsReference.PaymentServiceDoc.Links.budgetArticlesListLink.ToString()].Count > 0).
                Select(t => t[String.Format("[{0}]->[{1}]", Guids.RegCardsReference.PaymentServiceDoc.Links.budgetArticlesListLink.ToString(),
                Guids.RegCardsReference.PaymentServiceDoc.BudgetArticlesList.budgetArticleCode.ToString())]).ToList().Distinct().OrderBy(t => t.ToString());

            macro.ДиалогОжидания.Показать("Идет суммирование статей бюджета...", true);
            foreach (var articleCode in articleCodes)
            {
                rowCounter++;
                summaryRow = row + rowCounter;
                xls[7, summaryRow].NumberFormat = "#,##0.00 $";
                // Запись суммы в ячейку
                xls[7, summaryRow].SetFormula(String.Format("=СУММЕСЛИ(F6:F{0}; \"{1}\"; G6:G{0})", row, articleCode));

                xls[7, summaryRow].Font.Color = Color.Black.ToArgb();
                xls[7, summaryRow].Font.Bold = false;
                // Запись значения статьи бюджета в ячейку
                xls[6, summaryRow].SetValue(articleCode);
                // Границы ячеек кода статьи и суммы
                xls[6, summaryRow, 2, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight,
                                                                                                     XlBordersIndex.xlInsideVertical);
                // Добавление записи "Итого по коду:"
                xls[5, summaryRow].SetValue("Итого по коду:");

                double percent = articleCodes.Count() == 0 ? 0 : (double)rowCounter / articleCodes.Count() * 100;
                macro.ДиалогОжидания.СледующийШаг(String.Format("Суммируется статья {0}", articleCode), percent);

                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
            }
            macro.ДиалогОжидания.Скрыть();
            // Рамки ячеек до и после кода статьи и суммы
            xls[1, row + 1, 5, rowCounter].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                 XlBordersIndex.xlInsideHorizontal);
            xls[8, row + 1, 3, rowCounter].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                 XlBordersIndex.xlInsideHorizontal);
            return rowCounter;
        }
    }

    public class ReportHeaders
    {
        public int OrderDocHeader(Xls xls, DateTime startDate, DateTime finishDate)
        {
            int row = 2;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].Font.Bold = true;

            xls[1, row].SetValue("Реестр приказов по предприятию за период с " + startDate.ToShortDateString() + " по " + finishDate.ToShortDateString());
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 5, 1].Cells.MergeCells = true;

            row += 2;

            xls[1, row].SetValue("№ п/п");
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].ColumnWidth = 12.0;

            xls[2, row].SetValue("Дата");
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Size = 11.0;
            xls[2, row].ColumnWidth = 10.0;

            xls[3, row].SetValue("Наименование приказа");
            xls[3, row].Font.Bold = true;
            xls[3, row].Font.Size = 11.0;
            xls[3, row].ColumnWidth = 55.0;

            xls[4, row].SetValue("Кто подписал");
            xls[4, row].Font.Bold = true;
            xls[4, row].Font.Size = 11.0;
            xls[4, row].ColumnWidth = 30.0;

            xls[5, row].SetValue("Исполнитель");
            xls[5, row].Font.Bold = true;
            xls[5, row].Font.Size = 11.0;
            xls[5, row].ColumnWidth = 30.0;

            xls[1, row, 5, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, row, 5, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1, 5, row].Font.Name = "Arial";
            row++;
            return row;
        }
        public int InputDocHeader(Xls xls, DateTime startDate, DateTime finishDate)
        {
            int row = 2;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].Font.Bold = true;

            xls[1, row].SetValue("Реестр исходящих документов предприятия за период с " + startDate.ToShortDateString() + " по " + finishDate.ToShortDateString());
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 4, 1].Cells.MergeCells = true;

            row += 2;

            xls[1, row].SetValue("№ п/п");
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].ColumnWidth = 7.0;

            xls[2, row].SetValue("Адресат");
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Size = 11.0;
            xls[2, row].ColumnWidth = 50.0;

            xls[3, row].SetValue("Номер исходящего");
            xls[3, row].Font.Bold = true;
            xls[3, row].Font.Size = 11.0;
            xls[3, row].ColumnWidth = 20.0;

            xls[4, row].SetValue("Краткое содержание");
            xls[4, row].Font.Bold = true;
            xls[4, row].Font.Size = 11.0;
            xls[4, row].ColumnWidth = 60.0;

            xls[1, row, 4, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, 1, 4, row].Font.Name = "Arial";
            xls[1, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            row++;
            return row;
        }
        public int DecreeDocHeader(Xls xls, DateTime startDate, DateTime finishDate)
        {
            int row = 2;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].Font.Bold = true;

            xls[1, row].SetValue("Реестр распоряжений по предприятию за период с " + startDate.ToShortDateString() + " по " + finishDate.ToShortDateString());
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 5, 1].Cells.MergeCells = true;

            row += 2;

            xls[1, row].SetValue("№ п/п");
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].ColumnWidth = 12.0;

            xls[2, row].SetValue("Дата");
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Size = 11.0;
            xls[2, row].ColumnWidth = 10.0;

            xls[3, row].SetValue("Наименование распоряжения");
            xls[3, row].Font.Bold = true;
            xls[3, row].Font.Size = 11.0;
            xls[3, row].ColumnWidth = 55.0;

            xls[4, row].SetValue("Кто подписал");
            xls[4, row].Font.Bold = true;
            xls[4, row].Font.Size = 11.0;
            xls[4, row].ColumnWidth = 30.0;

            xls[5, row].SetValue("Исполнитель");
            xls[5, row].Font.Bold = true;
            xls[5, row].Font.Size = 11.0;
            xls[5, row].ColumnWidth = 30.0;

            xls[1, row, 5, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, row, 5, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1, 5, row].Font.Name = "Arial";

            row++;
            return row;
        }
        public int ContractDocHeader(Xls xls, DateTime startDate, DateTime finishDate)
        {
            int row = 2;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].Font.Bold = true;

            xls[1, row].SetValue("Реестр рамочных договоров за период с " + startDate.ToShortDateString() + " по " + finishDate.ToShortDateString());
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 6, 1].Cells.MergeCells = true;

            row += 2;

            xls[1, row].SetValue("От");
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].ColumnWidth = 10.0;

            xls[2, row].SetValue("Номер ДЗУ");
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Size = 11.0;
            xls[2, row].ColumnWidth = 30.0;

            xls[3, row].SetValue("Краткое содержание");
            xls[3, row].Font.Bold = true;
            xls[3, row].Font.Size = 11.0;
            xls[3, row].ColumnWidth = 60.0;

            xls[4, row].SetValue("Номер договора");
            xls[4, row].Font.Bold = true;
            xls[4, row].Font.Size = 11.0;
            xls[4, row].ColumnWidth = 30.0;

            xls[5, row].SetValue("Контрагент");
            xls[5, row].Font.Bold = true;
            xls[5, row].Font.Size = 11.0;
            xls[5, row].ColumnWidth = 30.0;

            xls[6, row].SetValue("Сумма с НДС");
            xls[6, row].Font.Bold = true;
            xls[6, row].Font.Size = 11.0;
            xls[6, row].ColumnWidth = 30.0;

            xls[1, row, 6, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, row, 6, 1].HorizontalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, 1, 6, row].Font.Name = "Arial";
            row++;

            return row;
        }
        public int ServiceDocHeader(Xls xls, DateTime startDate, DateTime finishDate, Объект department)
        {
            int row = 2;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].Font.Bold = true;
            xls[1, row, 7, 1].Cells.MergeCells = true;
            xls[1, row, 7, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row].SetValue("Реестр служебных записок подразделения " + department.ToString() + " за период с " + startDate.ToShortDateString() + " по " + finishDate.ToShortDateString());

            row += 2;

            xls[1, row].SetValue("№ С.З.");
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].ColumnWidth = 15.0;

            xls[2, row].SetValue("Дата");
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Size = 11.0;
            xls[2, row].ColumnWidth = 10.0;

            xls[3, row].SetValue("Исполнитель");
            xls[3, row].Font.Bold = true;
            xls[3, row].Font.Size = 11.0;
            xls[3, row].ColumnWidth = 35.0;

            xls[4, row].SetValue("Краткое содержание");
            xls[4, row].Font.Bold = true;
            xls[4, row].Font.Size = 11.0;
            xls[4, row].ColumnWidth = 30.0;

            xls[5, row].SetValue("Получатели");
            xls[5, row].Font.Bold = true;
            xls[5, row].Font.Size = 11.0;
            xls[5, row].ColumnWidth = 25.0;

            xls[6, row].SetValue("Визирующие");
            xls[6, row].Font.Bold = true;
            xls[6, row].Font.Size = 11.0;
            xls[6, row].ColumnWidth = 35.0;

            xls[7, row].SetValue("Стадия");
            xls[7, row].Font.Bold = true;
            xls[7, row].Font.Size = 11.0;
            xls[7, row].ColumnWidth = 12.0;

            xls[1, row, 7, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                          XlBordersIndex.xlEdgeBottom,
                                                                                          XlBordersIndex.xlEdgeLeft,
                                                                                          XlBordersIndex.xlEdgeRight);
            xls[1, 1, 7, row + 1].Font.Name = "Arial";

            row++;
            return row;
        }
        public int ReportDocHeader(Xls xls, DateTime startDate, DateTime finishDate, Объект department, bool incorrect)
        {
            int row = 2;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].Font.Bold = true;
            xls[1, row, 8, 1].Cells.MergeCells = true;
            xls[1, row, 8, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            if (!incorrect)
                xls[1, row].SetValue("Реестр докладных записок подразделения " + department.ToString() + " за период с " + startDate.ToShortDateString() + " по " + finishDate.ToShortDateString());
            else
                xls[1, row].SetValue("Реестр докладных записок без присоединенных файлов за период с " + startDate.ToShortDateString() + " по " + finishDate.ToShortDateString());

            row += 2;

            xls[1, row].SetValue("№ Д.З.");
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].ColumnWidth = 14.0;

            xls[2, row].SetValue("Дата");
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Size = 11.0;
            xls[2, row].ColumnWidth = 10.0;

            xls[3, row].SetValue("Исполнитель");
            xls[3, row].Font.Bold = true;
            xls[3, row].Font.Size = 11.0;
            xls[3, row].ColumnWidth = 35.0;

            xls[4, row].SetValue("Краткое содержание");
            xls[4, row].Font.Bold = true;
            xls[4, row].Font.Size = 11.0;
            xls[4, row].ColumnWidth = 40.0;

            xls[5, row].SetValue("Получатели");
            xls[5, row].Font.Bold = true;
            xls[5, row].Font.Size = 11.0;
            xls[5, row].ColumnWidth = 20.0;

            xls[6, row].SetValue("Подписал");
            xls[6, row].Font.Bold = true;
            xls[6, row].Font.Size = 11.0;
            xls[6, row].ColumnWidth = 35.0;

            xls[7, row].SetValue("Подр-е");
            xls[7, row].Font.Bold = true;
            xls[7, row].Font.Size = 11.0;
            xls[7, row].ColumnWidth = 10.0;

            xls[8, row].SetValue("Стадия");
            xls[8, row].Font.Bold = true;
            xls[8, row].Font.Size = 11.0;
            xls[8, row].ColumnWidth = 12.0;

            xls[1, row, 8, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                          XlBordersIndex.xlEdgeBottom,
                                                                                          XlBordersIndex.xlEdgeLeft,
                                                                                          XlBordersIndex.xlEdgeRight);
            xls[1, row, 8, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1, 8, row + 1].Font.Name = "Arial";
            row++;
            return row;
        }
        public int PaymentServiceDocHeader(Xls xls, DateTime startDate, DateTime finishDate, Объект department, bool isPaid)
        {

            int row = 2;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].Font.Bold = true;
            xls[1, row, 10, 1].Cells.MergeCells = true;
            xls[1, row, 10, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            string paidString = " ";
            if (isPaid)
                paidString = " оплаченных ";
            xls[1, row].SetValue(String.Format("Реестр{0}СЗНО/ЗПЗП подразделения {1} за период с {2} по {3}", paidString, department.ToString(), startDate.ToShortDateString(), finishDate.ToShortDateString()));

            row += 2;

            xls[1, row].SetValue("№ п/п");
            xls[1, row].Font.Bold = true;
            xls[1, row].Font.Size = 11.0;
            xls[1, row].ColumnWidth = 7.0;

            xls[2, row].SetValue("№ СЗНО");
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Size = 11.0;
            xls[2, row].ColumnWidth = 15.0;

            xls[3, row].SetValue("Дата");
            xls[3, row].Font.Bold = true;
            xls[3, row].Font.Size = 11.0;
            xls[3, row].ColumnWidth = 10.0;

            xls[4, row].SetValue("Контрагент");
            xls[4, row].Font.Bold = true;
            xls[4, row].Font.Size = 11.0;
            xls[4, row].ColumnWidth = 35.0;

            xls[5, row].SetValue("Краткое содержание");
            xls[5, row].Font.Bold = true;
            xls[5, row].Font.Size = 11.0;
            xls[5, row].ColumnWidth = 40.0;

            xls[6, row].SetValue("Код БДДС");
            xls[6, row].Font.Bold = true;
            xls[6, row].Font.Size = 11.0;
            xls[6, row].ColumnWidth = 30.0;

            xls[7, row].SetValue("Сумма");
            xls[7, row].Font.Bold = true;
            xls[7, row].Font.Size = 11.0;
            xls[7, row].ColumnWidth = 15.0;

            xls[8, row].SetValue("№ договора");
            xls[8, row].Font.Bold = true;
            xls[8, row].Font.Size = 11.0;
            xls[8, row].ColumnWidth = 15.0;

            xls[9, row].SetValue("№ заказа");
            xls[9, row].Font.Bold = true;
            xls[9, row].Font.Size = 11.0;
            xls[9, row].ColumnWidth = 15.0;

            xls[10, row].SetValue("Стадия");
            xls[10, row].Font.Bold = true;
            xls[10, row].Font.Size = 11.0;
            xls[10, row].ColumnWidth = 20.0;


            xls[1, row, 10, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                          XlBordersIndex.xlEdgeBottom,
                                                                                          XlBordersIndex.xlEdgeLeft,
                                                                                          XlBordersIndex.xlEdgeRight);

            xls[1, row, 10, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1, 10, row + 1].Font.Name = "Arial";
            row++;

            return row;
        }
    }
    public class ExcelOperations
    {
        // Параметры листа Excel (под формат А4)
        public void PageSetup(Xls xls)
        {
            PageSetup pageSetup = xls.Worksheet.PageSetup;
            pageSetup.PaperSize = XlPaperSize.xlPaperA4;
            pageSetup.Orientation = XlPageOrientation.xlLandscape;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 500;
            pageSetup.TopMargin = 0.5;
            pageSetup.BottomMargin = 0.5;
            pageSetup.LeftMargin = 1.0;
            pageSetup.RightMargin = 0.5;
            pageSetup.Zoom = false;
        }
    }
}
