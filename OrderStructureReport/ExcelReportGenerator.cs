using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ReportHelpers;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace OrderStructureReport
{
    public class ExcelReportGenerator
    {
        // Создание заголовка отчета
        public void CreateHeader(Xls xls)
        {
            xls.SelectWorksheet(2);
            xls.Worksheet.Name = "Структура заказа";
            xls[1, 2, 8, 2].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, 2, 8, 2].Font.Name = "Times New Roman";
            xls[1, 2, 8, 2].Font.Size = 12;
            xls[1, 2, 8, 2].Font.Bold = true;
            xls[1, 2, 8, 2].CenterText();
            xls[1, 2].Value = "№ п/п"; xls[1, 3].Value = "1";
            xls[2, 2].Value = "Шифр"; xls[2, 3].Value = "2";
            xls[3, 2].Value = "Наименование"; xls[3, 3].Value = "3";
            xls[4, 2].Value = "Обозначение"; xls[4, 3].Value = "4";
            xls[5, 2].Value = "Кол-во, шт"; xls[5, 3].Value = "5";
            xls[6, 2].Value = "Вид устройства"; xls[6, 3].Value = "6";
            xls[7, 2].Value = "Отдел-исполнитель"; xls[7, 3].Value = "7";
            xls[8, 2].Value = "Примечание"; xls[8, 3].Value = "8";
            xls[1, 2, 8, 2].WrapText = true;

            TableFormatter(xls);
        }

        // Создание тела отчета
        public void CreateReport(List<DocumentParameters> docList, Xls xls)
        {
            int row = 4;
            try
            {
                foreach (var item in docList)
                {
                    xls[1, row].NumberFormat = "@"; xls[1, row].Value = item.pointNumber;
                    xls[2, row].NumberFormat = "@"; xls[2, row].Value = item.UnitCode;
                    xls[3, row].Value = item.Name;
                    xls[4, row].Value = item.Denotation;
                    xls[5, row].Value = item.TotalAmount;
                    xls[6, row].Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, "=Классификатор", Type.Missing);
                    if (item.BomSection == "Прочие изделия")
                    {
                        xls[6, row].Value = "Покупное";
                        xls[7, row].Value = "Отдел 914";
                    }
                    else
                        xls[7, row].Value = "Отдел 20";
                    row++;
                }

                GroupTableRows(docList, xls);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

            xls[1, 4, 8, row - 4].Font.Name = "Times New Roman";
            xls[1, 4, 8, row - 4].Font.Size = 11;
            xls[1, 4, 8, row - 4].WrapText = true;
            xls[1, 4, 8, row - 4].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, 4, 2, row - 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[3, 4, 2, row - 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[5, 4, 3, row - 4].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[8, 4, 1, row - 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, 4, 8, row - 4].VerticalAlignment = XlVAlign.xlVAlignTop;
        }

        // Группировка строк таблицы отчета
        public void GroupTableRows(List<DocumentParameters> docList, Xls xls)
        {
            // Группировка сверху-вниз
            xls.Worksheet.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;

            try
            {
                for (int i = 1; i < docList.Count - 1; i++)
                {
                    int level = docList[i].Level;
                    int j = 1;

                    while ((i + j < docList.Count) && (docList[i + j].Level > level))
                    {
                        j++;
                    }
                    if (j > 1)
                    {
                        xls[1, 4 + i + 1, 1, j - 1].Rows.Group();
//                        xls[1, 4 + i + 1, 1, j - 1].EntireRow.Hidden = true;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        // Форматирование таблицы отчета
        private void TableFormatter (Xls xls)
        {
            xls[1, 1].Cells.ColumnWidth = 10.5;
            xls[2, 1].Cells.ColumnWidth = 13.5;
            xls[3, 1].Cells.ColumnWidth = 40.75;
            xls[4, 1].Cells.ColumnWidth = 19.75;
            xls[5, 1].Cells.ColumnWidth = 8.5;
            xls[6, 1].Cells.ColumnWidth = 13.5;
            xls[7, 1].Cells.ColumnWidth = 13.5;
            xls[8, 1].Cells.ColumnWidth = 14;

            PageSetup(xls, XlPaperSize.xlPaperA4, XlPageOrientation.xlLandscape, 28.5, 14.5, 14.5, 14.5);
        }

        // Настройки страницы Excel
        private void PageSetup(Xls xls, XlPaperSize paperSize, XlPageOrientation pageOrientation, double leftMargin, double rightMargin, double topMargin, double bottomMargin)
        {
            PageSetup pageSetup = xls.Worksheet.PageSetup;
            // Задать поля
            pageSetup.HeaderMargin = 0;
            pageSetup.LeftMargin = leftMargin;
            pageSetup.RightMargin = rightMargin;
            pageSetup.TopMargin = topMargin;
            pageSetup.BottomMargin = bottomMargin;
            // Отключение Zooma
            pageSetup.Zoom = false;
            // Разместить по ширине на одну страницу
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 500;
            // Размер бумаги, ориентацию страницы
            pageSetup.Orientation = pageOrientation;
            pageSetup.PaperSize = paperSize;
        }

        // Создание титульного листа
        public void CreateTitulList(string name, string denotation, string order, Xls xls)
        {
            xls.SelectWorksheet(1);
            xls.Worksheet.Name = "Титульный лист";

            xls[1, 1, 15, 34].Font.Name = "Times New Roman";
            xls[1, 1, 15, 34].Font.Size = 12;
            xls[1, 1, 15, 34].CenterText();

            xls[13, 1, 5, 1].Merge();
            xls[13, 1].Font.Bold = true;
            xls[13, 1].Value = "УТВЕРЖДАЮ";
            xls[13, 2, 5, 1].Merge();
            xls[13, 2].Value = "Генеральный директор";
            xls[13, 3, 5, 1].Merge();
            xls[13, 3].Value = "АО \"РКБ \"Глобус\"";
            xls[13, 5, 5, 1].Merge();
            xls[13, 5].Value = "_____________ Н.В.Гоев";
            xls[13, 7, 5, 1].Merge();
            xls[13, 7].Value = "\"___\"____________  " + DateTime.Now.Year;

            xls[4, 11, 11, 1].Merge();
            xls[4, 11].Font.Bold = true;
            xls[4, 11].Value = "СТРУКТУРА";

            xls[4, 13, 11, 1].Merge();
            xls[4, 14, 11, 1].Merge();
            xls[4, 15, 11, 1].Merge();
            xls[4, 13, 11, 3].Font.Bold = true;
            // Размещение наименования аппаратуры (должно быть не более трех строк)
            List<string> wrapName = WrapNameText("Аппаратуры \"" + name + "\"", 60);
            for (int i = 0; i < wrapName.Count; i++ )
            {
                xls[4, 13 + i].Value = wrapName[i];
            }

            xls[4, 17, 11, 1].Merge();
            xls[4, 17].Font.Bold = true;
            xls[4, 17].Value = denotation;

            xls[4, 19, 11, 1].Merge();
            xls[4, 19].Font.Bold = true;
            xls[4, 19].Value = "Заказ " + order;

            xls[1, 23, 5, 1].Merge();
            xls[1, 23].Value = "Начальник отдела 20";
            xls[1, 25, 5, 1].Merge();
            xls[1, 25].Value = "_____________ В.П.Зюзенков";
            xls[1, 27, 5, 1].Merge();
            xls[1, 27].Value = "\"___\"____________  " + DateTime.Now.Year;

            xls[13, 23, 5, 1].Merge();
            xls[13, 23].Value = "Главный конструктор";
            xls[13, 25, 5, 1].Merge();
            xls[13, 25].Value = "_____________ А.А.Мисник";
            xls[13, 27, 5, 1].Merge();
            xls[13, 27].Value = "\"___\"____________  " + DateTime.Now.Year;

            xls[1, 30, 5, 1].Merge();
            xls[1, 30].Value = "Начальник отдела 05";
            xls[1, 32, 5, 1].Merge();
            xls[1, 32].Value = "_____________ А.А.Трубников";
            xls[1, 34, 5, 1].Merge();
            xls[1, 34].Value = "\"___\"____________  " + DateTime.Now.Year;

            PageSetup(xls, XlPaperSize.xlPaperA4, XlPageOrientation.xlLandscape, 28.5, 10.5, 14.5, 14.5);
        }

        // Переносчик текстов
        private List<string> WrapNameText(string name, int sCount)
        {
            List<string> strings = new List<string>();
            // Количество строк, которое должен занять текст
            int lines = Convert.ToInt32(Math.Ceiling(((double)name.Length/(double)sCount)));
            char symbol;

            int nachIndex = 0;
            // Если количество строк больше одной, то разбиваем на несколько строк по пробелу или -, иначе передаем целиком
            if (lines > 1)
            {
                for (int i = 0; i < lines - 1; i++)
                {
                    int index = sCount + nachIndex;
                    symbol = name.ElementAt(index);
                    while ((symbol != ' ') && (symbol != '-'))
                    {
                        index--;
                        symbol = name.ElementAt(index);
                    }
                    strings.Add(name.Substring(nachIndex, index - nachIndex));
                    nachIndex = index + 1;
                }
                strings.Add(name.Substring(nachIndex, name.Length - nachIndex));
            }
            else
                strings.Add(name);

            return strings;
        }

        // Создание классификатора на третьем листе
        public void CreateClassificator(string[] items, Xls xls)
        {
            // Количество элементов в массиве items
            int cnt = items.Count();

            xls.SelectWorksheet(3);
            xls.Worksheet.Name = "Классификатор";

            xls[1, 1, 1, cnt].Font.Name = "Times New Roman";
            xls[1, 1, 1, cnt].Font.Size = 11;

            for (int i = 1; i <= cnt; i++)
                xls[1, i].Value = items[i - 1];

            xls[1, 1, 1, cnt].Select();
            xls.Application.ActiveWorkbook.Names.Add("Классификатор", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, "=" + xls.Worksheet.Name +"!R1C1:R" + cnt + "C1", Type.Missing);

            xls[1, 1].Cells.ColumnWidth = 27.5;
        }
    }

}
