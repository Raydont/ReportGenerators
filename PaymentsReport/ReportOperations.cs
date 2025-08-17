using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Desktop;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.IO;

using ReportHelpers;
using TFlex.DOCs.Model.Macros;

using System.Threading;
using TFlex.DOCs.Model.Macros.Flowchart.Activities;
using TFlex.DOCs.Model.Notification.ServerTasks.Triggers;
using TFlex.DOCs.Model.References.Links.Extensions;
using TFlex.DOCs.Model.Macros.ObjectModel;

namespace PaymentsReport
{
    public class ExcelOperations
    {
        // Настройки страницы Excel
        public void PageSetup(Xls xls, XlPaperSize paperSize, XlPageOrientation pageOrientation, int leftMargin, int rightMargin, int topMargin, int bottomMargin)
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
            pageSetup.FitToPagesTall = 100;
            // Размер бумаги, ориентацию страницы
            pageSetup.Orientation = pageOrientation;
            pageSetup.PaperSize = paperSize;

        }

        // Создание заголовка отчета
        public void CreateHeader(Xls xls)
        {
            string cell1 = String.Empty;
            string cell2 = String.Empty;

            int col = 1;
            int colStart = col;
            int row1 = 15;
            int row2 = 17;

            // Надписи в полях заголовка
            xls[col, row1].Cells.ColumnWidth = 5;
            col = InsertHeader("№ п/п", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 15;
            col = InsertHeader("Основание платежа", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 15;
            col = InsertHeader("№ платежного поручения", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 14;
            col = InsertHeader("Дата платежного поручения", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 15;
            col = InsertHeader("Контрагент", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 20;
            col = InsertHeader("Наименование товара,услуги", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 15;
            col = InsertHeader("Назначение платежа", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 20;
            col = InsertHeader("Вид оплаты", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 15;
            col = InsertHeader("Код БДДС", row1, col, "Calibri", 11, xls);

            xls[col, row1].Cells.ColumnWidth = 16;
            col = InsertHeader("Сумма к оплате", row1, col, "Calibri", 11, xls);

            col = InsertHeader("Справочно", row1, col, "Calibri", 11, xls);

            col = 11;

            xls[col, row2].Cells.ColumnWidth = 10;
            col = InsertHeader("Заказ", row2, col, "Calibri", 11, xls);

            xls[col, row2].Cells.ColumnWidth = 20;
            col = InsertHeader("Регистрационный № Договора", row2, col, "Calibri", 11, xls);

            xls[col, row2].Cells.ColumnWidth = 20;
            col = InsertHeader("№ Договора (оригинал)", row2, col, "Calibri", 11, xls);

            xls[col, row2].Cells.ColumnWidth = 10;
            col = InsertHeader("пункт ПТП", row2, col, "Calibri", 11, xls);

            xls[col, row2].Cells.ColumnWidth = 17;
            col = InsertHeader("Примечание", row2, col, "Calibri", 11, xls);

            xls[col, row2].Cells.ColumnWidth = 18;
            col = InsertHeader("Дата утверждения", row2, col, "Calibri", 11, xls);

            col = 1;

            // Запись номеров столбцов и их центрирование ячейках
            while (col < 17)
                col = InsertHeader(col.ToString(), row2 + 1, col, "Calibri", 11, xls);

            col = 1;

            // Объединение ячеек для первой половины заголовка
            for (int i = 0; i < 10; i++)
                xls[col + i, row1, 1, row2 - row1 + 1].Merge();

            col = 11; row1 = 15; row2 = 16;

            // Объединение ячеек для второй половины заголовка
            xls[col, row1, 6, row2 - row1 + 1].Merge();

            // Выделяем границы заголовка
            xls[1, row1, 16, row2 - row1 + 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                          XlBordersIndex.xlEdgeBottom,
                                                                                                          XlBordersIndex.xlEdgeLeft,
                                                                                                          XlBordersIndex.xlEdgeRight,
                                                                                                          XlBordersIndex.xlInsideVertical);
            xls[1, row2 + 1, 16, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                XlBordersIndex.xlEdgeBottom,
                                                                                                XlBordersIndex.xlEdgeLeft,
                                                                                                XlBordersIndex.xlEdgeRight,
                                                                                                XlBordersIndex.xlInsideVertical);
            // Закрашивание "шапки" таблицы отчета
            for (int i = 1; i <= 16; i++)
                for (int j = 15; j <= 18; j++)
                    xls[i, j].Interior.Color = ColorToInt(Color.LightSteelBlue);
        }

        // Преобразование цвета ARGB в целочисленное значение
        private int ColorToInt(Color color)
        {
            int colorInt = 0;

            colorInt = 255 << 24 | color.B << 16 | color.G << 8 | color.R;
            return colorInt;
        }

        // Заполнение таблицы Excel данными, полученными из выборки СЗНО
        public int FillTable(List<ReportParameters> paramList, Xls xls, MacroProvider macro)
        {
            int row = 19; int counter = 0;
            string cell = String.Empty;

            // открываем столбец "Примечание"
            xls[15, 19, 0, 500].Locked = false;
            xls[15, 19, 0, 500].FormulaHidden = false;

            // Границы ячеек таблицы
            xls[1, row, 16, paramList.Count].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            // Текстовый формат для всех колонок, кроме суммы (10-я колонка)
            xls[1, row, 9, paramList.Count].NumberFormat = "@";
            xls[10, row, 1, paramList.Count].NumberFormat = "#,##0.00 $";
            xls[11, row, 5, paramList.Count].NumberFormat = "@";

            macro.ДиалогОжидания.Показать("Идет заполнение таблицы Excel", true);
            foreach (ReportParameters parameters in paramList)
            {
                xls[1, row + counter].SetValue(parameters.number);
                xls[2, row + counter].SetValue(parameters.registrationNumber);
                if (parameters.isStagePaid)
                {
                    xls[3, row + counter].SetValue(parameters.draftNumber);
                    xls[4, row + counter].SetValue(parameters.draftDate);
                }
                xls[5, row + counter].SetValue(parameters.contractor);
                xls[6, row + counter].SetValue(parameters.productName);
                xls[7, row + counter].SetValue(parameters.paymentFunction);
                xls[8, row + counter].SetValue(parameters.paymentKind);
                xls[9, row + counter].SetValue(parameters.articleCode);
                xls[10, row + counter].SetValue(parameters.paymentSum);

                // Ячейка, в которую пишется сумма на оплату
                parameters.cell = Xls.Address(10, row + counter);

                xls[11, row + counter].SetValue(parameters.orderNumber);
                xls[12, row + counter].SetValue(parameters.contractRegNum);
                xls[13, row + counter].SetValue(parameters.contractNumber);
                xls[14, row + counter].SetValue(parameters.ptpPoint);
                xls[15, row + counter].SetValue(parameters.comment);

/*                if (parameters.isStagePaid)
                {
                    if ((parameters.draftDate == String.Empty) || (parameters.draftNumber == String.Empty))
                    {
                        xls[16, row + counter].Interior.Color = ColorToInt(Color.Red);
                        xls[16, row + counter].SetValue(parameters.stage + "\nНе указан номер ПП");
                    }
                    else
                        xls[16, row + counter].SetValue(parameters.stage + "\n" + parameters.draftNumber + "\nот " + parameters.draftDate);
                }
                else*/
                    xls[16, row + counter].SetValue(parameters.stage);

                counter++;
                double percent = paramList.Count == 0 ? 0 : (double)counter / paramList.Count * 100;
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                macro.ДиалогОжидания.СледующийШаг(String.Format("Записываются данные документа {0}", parameters.registrationNumber), percent);
            }
            macro.ДиалогОжидания.Скрыть();

            // Вычисление сумм по отдельным статьям бюджета, возвращает количество задействованных под статьи строк
            int rowCounter = ArticlesSumm(paramList, row + counter - 1, xls, macro);

            cell = Xls.Address(10, row);
            string endCell = Xls.Address(10, row + counter - 1);

            // Вычисление итоговой сумма к оплате (с учетом количества строк под суммы по статьям бюджета)
            string value = "=СУММ(" + cell + ":" + endCell + ")";
            int sumRow = row + counter + rowCounter;
            xls[10, sumRow].NumberFormat = "#,##0.00 $";
            xls[10, sumRow].SetFormula(value);
            xls[10, sumRow].Font.Color = Color.Black.ToArgb();
            xls[10, sumRow].Font.Bold = false;
            xls[10, sumRow].CenterText();
            string sumCell = Xls.Address(10, sumRow);

            // Запись "Итого платежей:"
            xls[7, row + counter + rowCounter].SetValue("Всего платежей:");
            xls[7, row + counter + rowCounter, 2, 1].Merge();
            xls[7, row + counter + rowCounter].CenterText();

            // Закрашивание строки "Всего платежей"
            for (int i = 1; i <= 16; i++)
                xls[i, row + counter + rowCounter].Interior.Color = Color.LightSteelBlue;

            // Границы ячеек кода статьи и суммы
            xls[9, row + counter + rowCounter, 2, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                                 XlBordersIndex.xlInsideVertical);
            // Рамки ячеек до и после кода статьи и суммы
            xls[1, row + counter + rowCounter, 8, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                                 XlBordersIndex.xlInsideHorizontal);

            xls[11, row + counter + rowCounter, 6, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                                 XlBordersIndex.xlInsideHorizontal);

            // Возвращаем ячейку "Всего платежей"
            return sumRow;
        }

        // Запись в ячейку элемента заголовка таблицы
        private int InsertHeader(string header, int row, int col, dynamic fontName, dynamic fontSize, Xls xls)
        {
            xls[col, row].Font.Name = fontName;
            xls[col, row].Font.Size = fontSize;
            xls[col, row].Font.Bold = true;
            xls[col, row].SetValue(header);
            col++;

            return col;
        }

        // Создание "шапки" документа (название документа, предприятия, период, подписи и т.д.)
        public void CreateDocTemplate(Xls xls, List<ReportParameters> paramList, int sumRow, string reportNumber, string reportDate, MacroProvider macro)
        {
            int col = InsertHeader("Реестр платежей и поступлений № " + reportNumber.Replace("РПП ","") + " от " + reportDate, 4, 1, "Arial", 16, xls);
            xls[1, 4, 10, 1].Merge();
            xls[1, 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, 4].VerticalAlignment = XlVAlign.xlVAlignCenter;

            string startDate = paramList[0].beginDate.ToShortDateString();
            string finishDate = paramList[0].endDate.ToShortDateString();
            col = InsertHeader("за период с " + startDate + " по " + finishDate, 6, 1, "Arial", 16, xls);
            xls[1, 6, 10, 1].Merge();
            xls[1, 6].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, 6].VerticalAlignment = XlVAlign.xlVAlignCenter;

            col = InsertHeader("Наименование общества", 9, 1, "Arial", 12, xls);
            xls[1, 9, 10, 1].Merge();
            xls[1, 9].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, 9].VerticalAlignment = XlVAlign.xlVAlignCenter;

            col = InsertHeader("ОАО \"Рязанское \"КБ Глобус\"", 10, 1, "Arial", 15, xls);
            xls[1, 10, 10, 1].Merge();
            xls[1, 10].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, 10].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, 10].Interior.Color = ColorToInt(Color.GreenYellow);

            col = InsertHeader("Платежи", 13, 1, "Arial", 16, xls);
            xls[1, 13, 10, 1].Merge();
            xls[1, 13].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, 13].VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Зам.ген.директора по экономике
            Объект post = macro.НайтиОбъект(Guids.Users.refGuid.ToString(), "[Guid]", Guids.Users.Objects.chiefEconomistGuid.ToString());
            string userPost = post[Guids.Users.Fields.userNameGuid.ToString()].ToString();
            Объект user = post.ДочерниеОбъекты.FirstOrDefault();
            string userName = user[Guids.Users.Fields.userShortNameGuid.ToString()].ToString();

            int row = sumRow;
            col = InsertHeader(userPost, row + 4, 3, "Arial", 13, xls);
            xls[3, row + 4, 5, 1].Merge();
            xls[3, row + 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[3, row + 4].VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Подчеркивание в месте подписи
            xls[10, row + 4, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeBottom);

            // ФИО зам.ген.директора
            col = InsertHeader(userName, row + 4, 11, "Arial", 13, xls);
            xls[11, row + 4, 3, 1].Merge();
            xls[11, row + 4].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[11, row + 4].VerticalAlignment = XlVAlign.xlVAlignCenter;
            // Вставка шейпа с подписью
/*            var pic = user[Guids.Users.Fields.signatureImageGuid].GetImage();
            SetPicture(10, row + 2, pic, xls, macro);*/

            // Для главного бухгалтера
            post = macro.НайтиОбъект(Guids.Users.refGuid.ToString(), "[Guid]", Guids.Users.Objects.chiefAccountGuid.ToString());
            userPost = post[Guids.Users.Fields.userNameGuid.ToString()].ToString();
            user = post.ДочерниеОбъекты.FirstOrDefault();
            userName = user[Guids.Users.Fields.userShortNameGuid.ToString()].ToString();

            col = InsertHeader(userPost, row + 8, 3, "Arial", 13, xls);
            xls[3, row + 8, 5, 1].Merge();
            xls[3, row + 8].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[3, row + 8].VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Подчеркивание в месте подписи
            xls[10, row + 8, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeBottom);

            // ФИО главбуха
            col = InsertHeader(userName, row + 8, 11, "Arial", 13, xls);
            xls[11, row + 8, 3, 1].Merge();
            xls[11, row + 8].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[11, row + 8].VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Вставка шейпа с подписью
/*            pic = user[Guids.Users.Fields.signatureImageGuid].GetImage();
            SetPicture(10, row + 6, pic, xls, macro);*/

            // Для начальника казначейства
            var dep = macro.НайтиОбъект(Guids.Users.refGuid.ToString(), "[Guid]", Guids.Users.Objects.dep101Guid.ToString());
            user = dep.СвязанныйОбъект[Guids.Users.Links.depChiefGuid.ToString()];
            userName = user[Guids.Users.Fields.userShortNameGuid.ToString()];

            col = InsertHeader("Начальник отдела казначейства", row + 12, 3, "Arial", 13, xls);
            xls[3, row + 12, 5, 1].Merge();
            xls[3, row + 12].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[3, row + 12].VerticalAlignment = XlVAlign.xlVAlignCenter;

            // Подчеркивание в месте подписи
            xls[10, row + 12, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeBottom);

            // ФИО начальника
            col = InsertHeader(userName, row + 12, 11, "Arial", 13, xls);
            xls[11, row + 12, 3, 1].Merge();
            xls[11, row + 12].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[11, row + 12].VerticalAlignment = XlVAlign.xlVAlignCenter;

            xls[1, 1, 2, 1].Merge();
        }

        // Вставка подписи
        [STAThread]
        private void SetPicture(int col, int row, Image img, Xls xls, MacroProvider macro)
        {
            try
            {
                // Вставка картинки из буфера обмена
                Thread thread = new Thread(() => ClipboardOperations(img, true));
                thread.SetApartmentState(ApartmentState.STA); 
                thread.Start();
                thread.Join();
                xls.Worksheet.Paste();

                xls.Worksheet.Shapes.Item(xls.Worksheet.Shapes.Count).Select();

                xls.Application.Selection.ShapeRange.Left = (float)xls[col, row].Left;
                xls.Application.Selection.ShapeRange.Top = (float)xls[col, row].Top;
                xls.Application.Selection.ShapeRange.Height = 49f;
                xls.Application.Selection.ShapeRange.Width = 85f;
                xls.Application.Selection.ShapeRange.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                xls.Application.Selection.ShapeRange.PictureFormat.TransparencyColor = System.Drawing.Color.White;
                xls.Application.Selection.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                // Очистка буфера обмена
                thread = new Thread(() => ClipboardOperations(img, false));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

            }
            catch (Exception ex)
            {
                macro.Сообщение("Ошибка вставки изображения", ex.ToString());
            }
        }

        private void ClipboardOperations(Image img, bool setImage)
        {
            Clipboard.Clear();
            if (setImage)
                Clipboard.SetImage(img);
        }

        private int ArticlesSumm(List<ReportParameters> paramList, int row, Xls xls, MacroProvider macro)
        {
            int summaryRow = 0;
            int rowCounter = 0;

            // Справочник "Статьи бюджета"
            var articleCodes = paramList.Where(t => !string.IsNullOrWhiteSpace(t.articleCode)).Select(t => t.articleCode).Distinct().OrderBy(t => t).ToList();

            macro.ДиалогОжидания.Показать("Идет суммирование статей бюджета...", true);
            foreach (var articleCode in articleCodes)
            {
                rowCounter++;
                summaryRow = row + rowCounter;
                xls[10, summaryRow].NumberFormat = "#,##0.00 $";
                // Запись суммы в ячейку
                xls[10, summaryRow].SetFormula(String.Format("=СУММЕСЛИ(I19:I{0}; \"{1}\"; J19:J{0})", row, articleCode));

                xls[10, summaryRow].Font.Color = Color.Black.ToArgb();
                xls[10, summaryRow].Font.Bold = false;
                xls[10, summaryRow].CenterText();
                // Запись значения статьи бюджета в ячейку
                xls[9, summaryRow].SetValue(articleCode);
                xls[9, summaryRow].CenterText();
                // Границы ячеек кода статьи и суммы
                xls[9, summaryRow, 2, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight,
                                                                                                     XlBordersIndex.xlInsideVertical);
                // Добавление записи "Итого по коду:"
                xls[7, summaryRow].SetValue("Итого по коду:");
                xls[7, summaryRow, 2, 1].Merge();
                xls[7, summaryRow].CenterText();
                // Закрашивание строки
                for (int i = 1; i <= 16; i++)
                    xls[i, summaryRow].Interior.Color = ColorToInt(Color.AliceBlue);

                double percent = articleCodes.Count == 0 ? 0 : (double)rowCounter / articleCodes.Count * 100;
                macro.ДиалогОжидания.СледующийШаг(String.Format("Суммируется статья {0}", articleCode), percent);

                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
            }
            macro.ДиалогОжидания.Скрыть();
            // Рамки ячеек до и после кода статьи и суммы
            xls[1, row + 1, 8, rowCounter].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                 XlBordersIndex.xlInsideHorizontal);
            xls[11, row + 1, 6, rowCounter].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                 XlBordersIndex.xlInsideHorizontal);
            return rowCounter;
        }
    }

    public class ObjectOperations
    {
        // Получение или создание папки
        public FolderObject GetChildFolder(FileReference fileReference, FolderObject parent, string path, string name)
        {
            FolderObject folderObject = (FolderObject)fileReference.FindByRelativePath(path);
            if (folderObject == null)
            {
                folderObject = parent.CreateFolder("", name);
                Desktop.CheckIn(folderObject, "Создание папки " + name, false);
            }

            return folderObject;
        }

        public string GetCorrectFileName(string fileName)
        {
            var newFileName = fileName;
            var invalidFileNameChars = new List<char>(Path.GetInvalidFileNameChars());
            invalidFileNameChars.ForEach(t => newFileName = newFileName.Replace(t, '_'));
            return newFileName.Trim();
        }
    }
    public class ReportParameters : ICloneable
    {
        // Поля отчета
        public DateTime beginDate { get; set; }                  // Начальная дата выборки СЗНО (сравнивается с параметром "От")
        public DateTime endDate { get; set; }                    // Конечная дата выбоки СЗНО (сравнивается с параметром "От")

        public string number { get; set; }                       // № п/п
        public string registrationNumber { get; set; }           // Основание платежа
        public string invoiceNumber { get; set; }                // Номер платежного документа
        public string invoiceData { get; set; }                  // Дата платежного документа
        public string contractor { get; set; }                   // Контрагент
        public string productName { get; set; }                  // Наименование товара, услуги
        public string paymentFunction { get; set; }              // Назначение платежа
        public string paymentKind { get; set; }                  // Вид оплаты
        public string articleCode { get; set; }                  // Код БДДС
        public double paymentSum { get; set; }                   // Сумма к оплате
        public string orderNumber { get; set; }                  // Номер заказа
        public string contractRegNum { get; set; }               // Регистрационный номер договора
        public string contractNumber { get; set; }               // Номер договора (оригинал)
        public string ptpPoint { get; set; }                     // Пункт ПТП
        public string comment { get; set; }                      // Примечание
        public string stage { get; set; }                        // Статус документа
        public string draftNumber { get; set; }                  // Номер платежного поручения
        public string draftDate { get; set; }                    // Дата платежного поручения
        public bool isStagePaid { get; set; }                    // Стадия - оплачено / не оплачено

        public string cell { get; set; }                         // Ячейка, в которую пишется сумма на оплату СЗНО

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }
}

    
