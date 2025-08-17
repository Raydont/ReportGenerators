using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model;
using TFlex.Drawing;
using TFlex.Model.Model2D;
using TFlex.Model;
using System.Windows.Forms.VisualStyles;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using TFlex.DOCs.Model.Macros.ObjectModel;

namespace AccountingCardReport
{
    public class ReportMaker
    {
        TFlex.Model.Document _cadDocument;
        public ReportMaker(TFlex.Model.Document cadDocument)
        {
            this._cadDocument = cadDocument;
        }
        // Получение таблицы документа
        private ParagraphText GetDocumentTextTable(string textTableName)
        {
            ParagraphText textTable = null;

            foreach (var text in _cadDocument.GetTexts())
            {
                if (text == null)
                    continue;

                if (text.Name == textTableName)
                    textTable = text as ParagraphText;

                if (textTable != null)
                    break;
            }

            if (textTable == null)
                throw new ApplicationException(
                    "На чертеже не найден объект " + textTableName);

            return textTable;
        }
        // Заполнение таблиц документа, наложение форматок и заполнение переменных
        public void FillDocument(DocumentParameters pars)
        {
            // Ориентировочное количество лицевых страниц (17 строк на одну страницу)
            var facePages = Math.Ceiling(pars.tableForm2.Count / 17.0) == 0 ? 1 : Convert.ToInt32(Math.Ceiling(pars.tableForm2.Count / 17.0));
            // Ориентировочное количество оборотных страниц (24 строки на одну страницу)
            var backPages = Math.Ceiling(pars.tableForm2a.Count / 24.0) == 0 ? 1 : Convert.ToInt32(Math.Ceiling(pars.tableForm2a.Count / 24.0));

            _cadDocument.BeginChanges("Заполнение таблиц и переменных");

            // Получение таблицы документа (Форма 2)
            ParagraphText textTable = GetDocumentTextTable("Form2Table");
            textTable.BeginEdit();
            Table table = textTable.GetTableByIndex(0);
            uint count = 0;
            var diff = backPages - facePages > 0 ? backPages - facePages : 0;
            // Количество добиваемых строк
            var strCount = pars.tableForm2.Count > 1 ? (uint)pars.tableForm2.Count - 1 + (uint)diff * 17 : (uint)diff * 17;

//            MessageBox.Show("Страниц - " + Math.Ceiling(strCount / 17.0));

            // Добивание пустыми строками
            table.InsertRows(strCount, 0, Table.InsertProperties.After);
            // Заполнение таблицы (Форма 2)
            foreach (var item in pars.tableForm2)
            {
                // Применяемость - Дата
                table.InsertText(4 + 10 * count, 0, item.Value[4]);
                // Применяемость - Обозначение
                if (item.Value[5].Contains("firstUsage"))
                {
                    // Если первичная применяемость, то подчеркивание
                    FormatCell(textTable, ref table, 5 + 10 * count, "T-FLEX Type T", 2.0, true, true, false);
                    table.InsertText(5 + 10 * count, 0, item.Value[5].Replace("firstUsage", ""), textTable.CharacterFormat);
                }
                else
                {
                    table.InsertText(5 + 10 * count, 0, item.Value[5]);
                }
                // Изменения - Номер изменения
                table.InsertText(6 + 10 * count, 0, item.Value[6]);
                // Изменения - Номер документа
                table.InsertText(7 + 10 * count, 0, item.Value[7]);
                // Изменения - Дата выпуска
                table.InsertText(8 + 10 * count, 0, item.Value[8]);
                // Изменения - Листы
                table.InsertText(9 + 10 * count, 0, item.Value[9]);
                count++;
            }
            textTable.EndEdit();

            // Получение таблицы документа (Форма 2а)
            textTable = GetDocumentTextTable("Form2aTable");
            textTable.BeginEdit();
            table = textTable.GetTableByIndex(0);
            diff = facePages - backPages > 0 ? facePages - backPages : 0;
            // Количество добиваемых строк
            strCount = pars.tableForm2a.Count > 1 ? (uint)pars.tableForm2a.Count - 1 + (uint)diff * 24 : (uint)diff * 24;
            // Добивание пустыми строками
            table.InsertRows(strCount, 0, Table.InsertProperties.After);
            // Заполнение таблицы (Форма 2a)
            count = 0;
            foreach (var item in pars.tableForm2a)
            {
                var tableString = item.Value;
                foreach (var str in tableString)
                {
                    if (!str.Contains("Header"))
                        FormatCell(textTable, ref table, count, "T-FLEX Type T", 2.0, false, false, false);
                    table.InsertText(count, 0, str.Replace("Header",""));
                    count++;
                }
            }
            // Объединение ячеек на оборотной стороне (для внешних абонентов)
            var bpCount = Convert.ToInt32(Math.Ceiling(pars.tableForm2a.Count / 24.0));
            if (bpCount > 0)
            {
                for (int i = 0; i < bpCount; i++)
                {
                    for (int j = 0; j < 11; j++)
                    {
                        var startCell = (uint)(j + (i * 23 * 11));
                        var endCell = (uint)((i * 23 * 11) + 11);
                        table.MergeCells(startCell, endCell);
                    }
                }
            }
            textTable.EndEdit();

            var pagesCollection = _cadDocument.GetPages();
            var pagesCount = pagesCollection.Count;                                               // Всего страниц
            var backPage = pagesCollection.First(t => t.Name == "Обратная").Rank;                 // Порядковый номер первой оборотной страницы
            var pageNumber = 1;                                                                   // Номер лицевой страницы

            // Наложение форматок и заполнение переменных
            var fileReference = new FileReference(ServerGateway.Connection);
            var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.Form2Gost2_501_2013);
            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
            var pageForm2Link = new FileLink(_cadDocument, fileObject.LocalPath, true);
            fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.Form2aGost2_501_2013);
            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
            var pageForm2aLink = new FileLink(_cadDocument, fileObject.LocalPath, true);

            int pageCounter = 0;
            foreach (var page in pagesCollection)
            {
                if (page.Rank == backPage)
                {
                    pageCounter = 0;
                }
                pageCounter++;

                if (page.Rank < backPage)
                {
                    page.Name = "Лицевая " + pageCounter;
                    page.Rectangle = new Rectangle(0, 0, 210, 148);

                    Fragment formFragment = new Fragment(pageForm2Link);
                    formFragment.Page = page;

                    // задание значений переменных фрагмента
                    foreach (FragmentVariableValue variable in formFragment.GetVariables())
                    {
                        switch (variable.Name)
                        {
                            case "$bfmi":
                                variable.TextValue = pars.manufactoryCode; break;
                            case "$dep":
                                variable.TextValue = pars.department; break;
                            case "$doctype":
                                variable.TextValue = pars.doctype; break;
                            case "$name":
                                variable.TextValue = pars.name; break;
                            case "$invnum":
                                variable.TextValue = pars.invnum; break;
                            case "$denotation":
                                variable.TextValue = pars.denotation; break;
                            case "$format":
                                variable.TextValue = pars.format; break;
                            case "$indate":
                                variable.TextValue = pars.incomingDate; break;
                            case "$cpages":
                                variable.TextValue = pars.sheetsCount; break;
                            case "$list":
                                variable.TextValue = pageNumber.ToString(); break;
                            case "$next":
                                variable.TextValue = pageNumber + 1 > backPage ? String.Empty : (pageNumber + 1).ToString(); break;
                        }
                    }
                    pageNumber++;
                }
                else
                {
                    page.Name = "Обратная " + pageCounter;
                    page.Rectangle = new Rectangle(0, 0, 210, 148);

                    Fragment formFragment = new Fragment(pageForm2aLink);
                    formFragment.Page = page;
                }
            }
            _cadDocument.EndChanges();
        }

        public static void FormatCell(ParagraphText parText, ref Table parTextTable, uint cell, string fontName, double fontSize, bool underlined, bool bold, bool italic)
        {
            CharFormat newCharacterFormat = parText.CharacterFormat;
            newCharacterFormat.isBold = bold;
            newCharacterFormat.isUnderline = underlined;
            newCharacterFormat.isItalic = italic;
            newCharacterFormat.isDefaultItalic = false;
            newCharacterFormat.Color = parText.CharacterFormat.Color;
            newCharacterFormat.FontName = fontName;
            newCharacterFormat.FontSize = fontSize;
            newCharacterFormat.isStrikeout = parText.CharacterFormat.isStrikeout;
            newCharacterFormat.Space = parText.CharacterFormat.Space;
            newCharacterFormat.VertOffset = parText.CharacterFormat.VertOffset;

            parTextTable.SetSelection(cell, cell);
            parText.CharacterFormat = newCharacterFormat;
        }
    }
}
