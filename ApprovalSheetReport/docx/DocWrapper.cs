using System;
using Microsoft.Office.Interop.Word;

namespace ReportHelpers
{
    public class WordDoc : IDisposable
    {
        private Application _application;
        public Application Application
        {
            get { return _application ?? (_application = new Application { Visible = _visible }); }
        }

        private Document _document;
        public Document Document
        {
            get { return _document ?? (_document = Application.Documents.Add()); }
        }

        private bool _visible;
        public bool Visible
        {
            get
            {
                return _visible;
            }
            set
            {

                _visible = value;
                if (_application != null)
                {
                    _application.Visible = _visible;
                }
            }
        }

        #region Общие настройки

        public void PageSetup(WdOrientation orientation)
        {
            Document.PageSetup.Orientation = orientation;
        }

        public void PageSetup(WdOrientation orientation, float marginLeft, float marginRight)
        {
            Document.PageSetup.Orientation = orientation;
            Document.PageSetup.LeftMargin = CentimetersToPoints(marginLeft);
            Document.PageSetup.RightMargin = CentimetersToPoints(marginRight);
        }

        #endregion

        #region Работа с документами

        public void Open(string docPath, bool saveLast = true, bool readOnly = false)
        {
            if (_document != null)
            {
                _document.Close(saveLast);
            }
            _document = null;
            _document = Application.Documents.Open(docPath, ReadOnly: readOnly);
        }

        public void Close(bool saveChanges = false)
        {
            if (_document != null)
            {
                _document.Close(saveChanges);
                _document = null;
            }
        }

        public void Save()
        {
            if (_document != null)
            {
                _document.Save();
            }
        }

        public void SaveAs(string docPath)
        {
            if (_document != null) _document.SaveAs(docPath);
        }

        public void Dispose()
        {
            Quit();
        }

        public void Quit(bool saveChanges = false)
        {
            if (_application == null) return;

            Close(saveChanges);

            _application.Quit();
            ClearApplicationLink();
            _application = null;
        }

        public void ClearApplicationLink()
        {
            while (System.Runtime.InteropServices.Marshal.ReleaseComObject(_application) > 0)
            {
            }
            GC.Collect();
            _application = null;
        }

        #endregion

        #region Работа с таблицами

        public Table CreateTable(Range range, int numRows, int numColumns, WdAutoFitBehavior autofit)
        {
            return Document.Tables.Add(range, numRows, numColumns, WdDefaultTableBehavior.wdWord9TableBehavior, autofit);
        }

        public Range GetRange(Cell cellStart, Cell cellEnd)
        {
            return Document.Range(cellStart.Range.Start, cellEnd.Range.End);
        }

        #endregion

        #region Работа с текстом

        public static void StyleBold(object rangeObject)
        {
            var range = (Range)rangeObject;
            range.Font.Bold = 1;
        }

        public static void StyleItalic(object rangeObject)
        {
            var range = (Range)rangeObject;
            range.Font.Italic = 1;
        }

        public static void StyleUnderline(object rangeObject)
        {
            var range = (Range)rangeObject;
            range.Font.Underline = WdUnderline.wdUnderlineSingle;
        }

        public static void StyleCenterTextHorizontal(object rangeObject)
        {
            var range = (Range)rangeObject;
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        public static void StyleLeftTextHorizontal(object rangeObject)
        {
            var range = (Range)rangeObject;
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        }

        public static void StyleCenterTextVertical(object rangeObject)
        {
            var range = (Range)rangeObject;
            range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;           
        }

        public static void StyleBackgroundGray(object rangeObject)
		{
            var range = (Range)rangeObject;
            range.Cells.Shading.BackgroundPatternColor = WdColor.wdColorGray20;
        }

        /// <summary>
        /// Вставить параграф
        /// </summary>
        /// <example>
        /// <code>wordDoc.InsertParagraph(rng, signBlock.GetSignTypeString(), WdParagraphAlignment.wdAlignParagraphCenter,
        /// rangeObject =>
        /// {
        ///    var range = (Range)rangeObject;
        ///    range.Font.Bold = 1;
        ///    range.Paragraphs[1].CharacterUnitLeftIndent = 1;
        /// });
        /// </code>
        /// </example>
        /// <param name="rng">диапазон</param>
        /// <param name="text">текст параграфа</param>
        /// <param name="alignParagraph">выравнивание в параграфе</param>
        public void InsertParagraph(Range rng, string text, WdParagraphAlignment alignParagraph, params Action<object>[] editParagraphs)
        {
            if (rng == null)
            {
                rng = Document.Range();
            }
            var countParagraphs = rng.Paragraphs.Count;
            rng.InsertAfter(text);
            rng.InsertParagraphAfter();

            for (int i = 0; i < rng.Paragraphs.Count - countParagraphs; i++)
            {
                var paragraph = rng.Paragraphs[rng.Paragraphs.Count - 1 - i];
                paragraph.Alignment = alignParagraph;

                foreach (var editParagraph in editParagraphs)
                {
                    if (editParagraph != null)
                    {
                        editParagraph(rng.Paragraphs[rng.Paragraphs.Count -1 - i].Range);
                    }
                }
            }
        }

        public static void VerticalText(object rangeObject)
        {
            var range = (Range)rangeObject;
            range.Orientation = WdTextOrientation.wdTextOrientationUpward;
        }

        #endregion

        #region Конвертация единиц

        public float CentimetersToPoints(float centimeters)
        {
            return Application.CentimetersToPoints(centimeters);
        }

        public float InchesToPoints(float inches)
        {
            return Application.InchesToPoints(inches);
        }

        public float MillimetersToPoints(float millimeters)
        {
            return Application.MillimetersToPoints(millimeters);
        }

        public float PicasToPoints(float picas)
        {
            return Application.PicasToPoints(picas);
        }

        #endregion

        public bool SetShapeValue(string key, string text)
        {
            var result = false;
            foreach (Shape shape in Document.Shapes)
            {
                if (string.IsNullOrWhiteSpace(shape.AlternativeText)) continue;

                var shapeKey = shape.AlternativeText.ToUpper();

                if (shapeKey == key.ToUpper())
                {
                    shape.TextFrame.TextRange.Text = text;
                    result = true;
                }
            }
            return result;
        }

        public bool SetBookmarkValue(string key, string text)
        {
            var result = false;
            foreach (Bookmark bookmark in Document.Bookmarks)
            {
                if (bookmark.Name.ToUpper() == key)
                {
                    bookmark.Range.Text = text;
                    result = true;
                }
            }
            return result;
        }

        public bool SetTextFieldValue(string key, string text)
        {
            var result = false;
            foreach (FormField formField in Document.FormFields)
            {
                if (formField.Type != WdFieldType.wdFieldFormTextInput) continue;
                if (formField.Name == key)
                {
                    formField.Result = text;
                    result = true;
                }
            }
            return result;
        }

        public bool SetReportValue(string key, string text)
        {
            return SetShapeValue(key, text) | SetBookmarkValue(key, text) | SetTextFieldValue(key, text);
        }

        public Bookmark GetBookmark(string key)
        {
            foreach (Bookmark bookmark in Document.Bookmarks)
            {
                if (bookmark.Name.ToUpper() == key)
                {
                    return bookmark;
                }
            }
            return null;
        }
    }

    public static class DocExtender
    {
        public static void SetWidthCentimeters(this Column column, float centimeters)
        {
            column.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
            column.PreferredWidth = column.Application.CentimetersToPoints(centimeters);
        }

        public static void SetWidthPercent(this Column column, float percent)
        {
            column.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            column.PreferredWidth = percent;
        }

        public static void SetText(this Table table, int col, int row, string text)
        {
            table.Rows[row].Cells[col].Range.Text = text;
        }

        public static void CenterText(this Range range)
        {
            range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        public static void Bold(this Range range, bool bold = true)
        {
            range.Font.Bold = bold ? 1 : 0;
        }
    }
}


