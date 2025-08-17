using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogisticsControlReport
{
    public delegate void ProcessCellDelegate(ref XlsxCell cell);

    public class XlsxDocument
    {
        public string fontName = "Calibri";
        public double fontSize = 11;
        public System.Drawing.Color fontColor = System.Drawing.Color.Black;

        public readonly List<XlsxSheet> sheets = new List<XlsxSheet>();

        int currentSheetIndex = -1;


        public XlsxDocument()
        {
            EnshureWorksheetCount(1);
        }

        public XlsxSheet Worksheet
        {
            get
            {
                if (currentSheetIndex < 0)
                {
                    if (sheets.Count == 0)
                    {
                        sheets.Add(new XlsxSheet(this));
                    }
                    currentSheetIndex = 0;
                }
                return sheets[currentSheetIndex];
            }
            set
            {
                currentSheetIndex = sheets.IndexOf(value);
            }
        }

        public XlsxSheet AddSheet(string name, bool setCurrent = true)
        {
            var sheet = new XlsxSheet(this, name);
            sheets.Add(sheet);
            if (setCurrent) sheet.SetCurrent();
            return sheet;
        }

        public void EnshureWorksheetCount(int count)
        {
            while (sheets.Count < count)
            {
                AddSheet("Лист" + (sheets.Count + 1));
            }
        }

        public void SelectWorksheet(int index)
        {
            sheets[index - 1].SetCurrent();
        }

        public XlsxRange this[int x, int y]
        {
            get
            {
                return new XlsxRange(Worksheet, x, y, 1, 1);
            }
        }

        public XlsxRange this[int x, int y, int width, int height]
        {
            get
            {
                return new XlsxRange(Worksheet, x, y, width, height);
            }
        }


        public void GenerateFile(string filePath)
        {
            var generator = new XlsxSpreadsheetGenerator();
            generator.GenerateFile(filePath, this);
        }
        public byte[] GenerateFileInMemory()
        {
            var generator = new XlsxSpreadsheetGenerator();
            return generator.GenerateToMemory(this);
        }

        public void AutoWidth()
        {
            var columnsCount = Worksheet.columnsCount;
            for (int i = 0; i < columnsCount; i++)
            {
                var column = Worksheet.GetColumn(i + 1);

                if (Worksheet.ColumHasData(i + 1))
                    column.columnWidth = Worksheet.GetMaxColumnWidth(i + 1);
            }
        }

        public void SaveToInMyDocumentsAndOpen(string fileName)
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\";

            int index = 0;

            while (true)
            {
                var filePath = Path.Combine(documents, fileName + (index > 0 ? index.ToString() : string.Empty) + ".xlsx");
                bool canCreate = true;
                if (File.Exists(filePath))
                {
                    try
                    {
                        File.Delete(filePath);
                    }
                    catch
                    {
                        canCreate = false;
                    }
                }

                if (canCreate)
                {
                    GenerateFile(filePath);
                    Process.Start(filePath);
                    break;
                }
                else
                {
                    index++;
                }
            }

        }

        public void SaveTo(string filePath)
        {
            GenerateFile(filePath);
        }
    }

    public class XlsxSheet
    {
        public readonly XlsxDocument document;

        public string Name;

        public double defaultRowHeight = 15;
        public double defaultColumnWidth = 10;
        public string AutoFilterRange;

        List<XlsxColumn> columns = new List<XlsxColumn>();
        List<XlsxRow> rows = new List<XlsxRow>();

        public int rowsCount => rows.Count;
        public int columnsCount;
        public int columnsCountWithData
        {
            get
            {
                int maxCol = 1;

                for (int i = 0; i < rows.Count; i++)
                {
                    XlsxRow row = rows[i];
                    for (int j = 0; j < row.cellsCount; j++)
                    {
                        row.ProcessCell(j + 1, (ref XlsxCell cell) =>
                        {
                            if (cell.value != null || cell.formula != null)
                            {
                                maxCol = Math.Max(j + 1, maxCol);
                            }
                        }, false);
                    }
                }

                return maxCol;
            }
        }

        XlsxOutline _Outline;
        public XlsxOutline Outline
        {
            get
            {
                if (_Outline == null) _Outline = new XlsxOutline(this);
                return _Outline;
            }
        }

        public XlsxPageSetup PageSetup = new XlsxPageSetup();

        public int SplitColumn { get; internal set; }
        public int SplitRow { get; internal set; }
        public bool FreezePanes { get; internal set; }

        public XlsxSheet(XlsxDocument document, string name = "Лист 1")
        {
            this.document = document;
            this.Name = name;
        }

        public XlsxSheet SetCurrent()
        {
            document.Worksheet = this;
            return this;
        }

        public XlsxColumn GetColumn(int x)
        {
            while (columns.Count < x)
            {
                columns.Add(new XlsxColumn(this, x));
            }

            return columns[x - 1];
        }

        public XlsxRow GetRow(int y)
        {
            while (rows.Count < y)
            {
                var rowY = rows.Count + 1;
                rows.Add(new XlsxRow(this, rowY));
            }

            return rows[y - 1];
        }

        public double GetMaxColumnWidth(int column)
        {
            double maxLength = 0;

            for (int i = 0; i < Math.Min(rowsCount, 1000); i++)
            {
                string value = string.Empty;
                XlsxCell cell = default;

                ProcessCell(column, i + 1, (ref XlsxCell c) =>
                {
                    cell = c;
                });

                if (cell.colSpan > 1) continue;

                if (cell.value is XlsxDecoratedText)
                {
                    value = (cell.value as XlsxDecoratedText).GetText();
                }
                else
                {
                    value = cell.value is string ? (string)cell.value : cell.value + string.Empty;
                }

                var font = GetFont(cell);

                if (font != null)
                {
                    var length = System.Windows.Forms.TextRenderer.MeasureText(value, font).Width / 7d + 2;
                    maxLength = Math.Max(maxLength, length);
                }
                else
                {


                    var maxLineLength = 0;
                    var currentLength = 0;
                    for (int i1 = 0; i1 < value.Length; i1++)
                    {
                        char ch = value[i1];

                        if (ch != '\r' && ch != '\n')
                        {
                            currentLength += 1;
                        }

                        if (ch == '\n' || i1 == value.Length - 1)
                        {
                            if (currentLength > maxLineLength) maxLineLength = currentLength;
                            currentLength = 0;
                        }

                    }


                    var length = 15;

                    //var lines = value.Split('\n');
                    //if (lines.Length > 0)
                    //{
                    //    length = lines.Max(t => t.Length);
                    //}

                    if (maxLineLength > 0) length = maxLineLength;

                    maxLength = Math.Max(maxLength, length * (cell.bold ? 1.2 : 1) * ((11d + (cell.fontSize - 11) * 4) / 11d));
                }
            }
            return Math.Min(255, maxLength);
        }

        internal bool ColumHasData(int column)
        {
            for (int i = 0; i < rowsCount; i++)
            {
                object value = null;
                ProcessCell(column, i + 1, (ref XlsxCell cell) =>
                {
                    value = cell.value;
                });

                if (value != null && !string.IsNullOrWhiteSpace(value.ToString())) return true;
            }

            return false;
        }

        //Dictionary<string, System.Drawing.Font> fonts = new Dictionary<string, System.Drawing.Font>();

        Dictionary<string, Dictionary<double, System.Drawing.Font>> fonts = new Dictionary<string, Dictionary<double, System.Drawing.Font>>();

        Dictionary<string, XlsxComment> comments = new Dictionary<string, XlsxComment>();

        public bool hasComments => comments.Count > 0;

        public XlsxComment GetComment(string adress)
        {
            comments.TryGetValue(adress, out XlsxComment comment);
            return comment;
        }

        public XlsxComment SetComment(string address, string text)
        {
            comments.TryGetValue(address, out XlsxComment comment);
            if (comment == null)
            {
                comment = new XlsxComment();
                comment.address = address;
                comments[address] = comment;
            }
            comment.comment = text;
            return comment;
        }


        public List<XlsxComment> GetComments()
        {
            return comments.Values.ToList();
        }

        private System.Drawing.Font GetFont(XlsxCell cell)
        {
            //var fontKey = cell.fontName + cell.fontSize.ToString("f1");
            System.Drawing.Font font;

            if (fonts.TryGetValue(cell.fontName, out Dictionary<double, System.Drawing.Font> fontsBySize))
            {
                if (fontsBySize.TryGetValue(cell.fontSize, out font))
                {
                    return font;
                }
            }
            else
            {
                fonts[cell.fontName] = new Dictionary<double, System.Drawing.Font>();
            }

            font = new System.Drawing.Font(cell.fontName, (float)cell.fontSize, cell.bold ? System.Drawing.FontStyle.Bold : System.Drawing.FontStyle.Regular);

            fonts[cell.fontName][cell.fontSize] = font;

            return font;
        }


        public void ProcessCell(int x, int y, ProcessCellDelegate action, bool create = true)
        {
            if (create)
            {
                while (rows.Count < y)
                {
                    var rowY = rows.Count + 1;
                    rows.Add(new XlsxRow(this, rowY));
                }
            }

            if (rows.Count < y) return;

            rows[y - 1].ProcessCell(x, (ref XlsxCell cell) =>
            {
                action(ref cell);
            }, create);
        }
    }

    public class XlsxComment
    {
        public string address;
        public string comment;
    }

    public class XlsxPageSetup
    {
        public XlsxPageSetup()
        {
        }

        public OrientationValues orientation = OrientationValues.Portrait;
        public bool? Zoom = false;
        public uint? FitToPagesWide = 1;
        public uint? FitToPagesTall = 10000;
    }

    public enum XlsxSummaryRow
    {
        Above = 0,
        Below = 1
    }

    public enum XlsxSummaryColumn
    {
        Left = 0,
        Right = 1
    }

    public class XlsxOutline
    {
        XlsxSheet sheet;

        public XlsxSummaryRow SummaryRow;
        public XlsxSummaryColumn SummaryColumn;

        public int RowLevels = 0;
        public int ColumnLevels = 0;

        public XlsxOutline(XlsxSheet sheet)
        {
            this.sheet = sheet;
        }

        public void ShowLevels(int? RowLevels = null, int? ColumnLevels = null)
        {
            if (RowLevels.HasValue) this.RowLevels = RowLevels.Value;
            if (ColumnLevels.HasValue) this.ColumnLevels = ColumnLevels.Value;
        }
    }

    public class XlsxRow
    {
        public readonly XlsxSheet sheet;
        public readonly int rowIndex;
        public double? rowHeight;
        public int outlineLevel;

        List<XlsxCell> cells;


        public int cellsCount => cells.Count;

        public XlsxRow(XlsxSheet sheet, int rowIndex)
        {
            this.sheet = sheet;
            this.rowIndex = rowIndex;
            int firstRowSize = 0;
            if (sheet.rowsCount > 0)
                firstRowSize = sheet.GetRow(1).cells.Count + 5;
            cells = new List<XlsxCell>(Math.Max(32, firstRowSize));
        }

        public void ProcessCell(int x, ProcessCellDelegate action, bool create = true)
        {
            if (create)
            {
                while (cells.Count < x)
                {
                    var cellX = cells.Count + 1;
                    cells.Add(new XlsxCell(this, cellX));
                }

            }
            if (cells.Count < x) return;

            var cell = cells[x - 1];
            action(ref cell);
            cells[x - 1] = cell;
        }

        //*
        //public XlsxCell this[int x]
        //{
        //    get
        //    {
        //        while (cells.Count < x)
        //        {
        //            cells.Add(new XlsxCell(this, x, rowIndex));
        //        }
        //        return cells[x - 1];
        //    }
        //    set
        //    {

        //    }
        //}

    }

    public struct XlsxCell
    {
        public readonly XlsxRow row;
        public readonly int colIndex;
        public int rowIndex => row.rowIndex;

        public HorizontalAlignmentValues hAlign;
        public VerticalAlignmentValues vAlign;

        internal int styleIndex;

        public bool italic;
        public bool underline;
        public bool bold;


        public object value;
        public string formula;
        public string numberFormat;

        public XlsxCellBorders cellBorders;

        string _fontName;
        double? _fontSize;
        System.Drawing.Color? _fontColor;

        public System.Drawing.Color? interiorColor;

        public bool wrapText;
        public int colSpan;
        public int rowSpan;

        public int mergedToX;
        public int mergedToY;
        public XlsxCell(XlsxRow row, int colIndex)
        {
            this.row = row;
            this.colIndex = colIndex;

            if (row.sheet.columnsCount < colIndex)
            {
                row.sheet.columnsCount = colIndex;
            }
            hAlign = HorizontalAlignmentValues.General;
            vAlign = VerticalAlignmentValues.Bottom;

            wrapText = false;
            colSpan = 1;
            rowSpan = 1;
            styleIndex = 0;
            italic = false;
            underline = false;
            bold = false;
            value = null;
            formula = null;
            numberFormat = null;
            cellBorders = new XlsxCellBorders();
            _fontName = null;
            _fontSize = null;
            _fontColor = null;
            interiorColor = null;
            mergedToX = -1;
            mergedToY = -1;
        }



        public string fontName
        {
            get
            {
                return _fontName ?? row.sheet.document.fontName;
            }
            set
            {
                _fontName = value;
            }
        }

        public double fontSize
        {
            get
            {
                return _fontSize ?? row.sheet.document.fontSize;
            }
            set
            {
                _fontSize = value;
            }
        }

        public System.Drawing.Color fontColor
        {
            get
            {
                return _fontColor ?? row.sheet.document.fontColor;
            }
            set
            {
                _fontColor = value;
            }
        }
    }

    public class XlsxCellInteriorRange
    {
        XlsxRange range;
        public XlsxCellInteriorRange(XlsxRange range)
        {
            this.range = range;
        }


        public System.Drawing.Color? Color
        {
            get
            {
                return range.InteriorColor;
            }
            set
            {
                range.InteriorColor = value;
            }
        }
    }


    public enum XlsxBorder : byte
    {
        Top = 0,
        Left,
        Right,
        Bottom,
        InsideVertical,
        InsideHorizontal
    }

    public struct XlsxCellBorders
    {
        XlsxCellBorder top;
        XlsxCellBorder left;
        XlsxCellBorder right;
        XlsxCellBorder bottom;

        public BorderStyleValues GetBorderWidth(XlsxBorder border)
        {
            switch (border)
            {
                case XlsxBorder.Top: return top.width;
                case XlsxBorder.Left: return left.width;
                case XlsxBorder.Right: return right.width;
                case XlsxBorder.Bottom: return bottom.width;
            }

            return BorderStyleValues.None;
        }

        public void SetBorderWidth(XlsxBorder border, BorderStyleValues width)
        {
            switch (border)
            {
                case XlsxBorder.Top: top.width = width; return;
                case XlsxBorder.Left: left.width = width; return;
                case XlsxBorder.Right: right.width = width; return;
                case XlsxBorder.Bottom: bottom.width = width; return;
            }
        }
    }

    public struct XlsxCellBorder
    {
        public BorderStyleValues width;

        public XlsxCellBorder(BorderStyleValues width = BorderStyleValues.None)
        {
            this.width = width;
        }
    }

    public class XlsxRange
    {
        public readonly XlsxSheet sheet;
        public int x;
        public int y;
        public int width;
        public int height;

        public XlsxRange(XlsxSheet sheet, int x, int y, int width, int height)
        {
            this.sheet = sheet;
            this.x = x;
            this.y = y;
            this.width = width;
            this.height = height;
        }

        public XlsxFontRange Font => new XlsxFontRange(this);

        public XlsxColumnRange EntireColumn => new XlsxColumnRange(this, true);
        public XlsxColumnRange Columns => new XlsxColumnRange(this, false);
        public XlsxRowsRange EntireRow => new XlsxRowsRange(this, true);
        public XlsxRowsRange Rows => new XlsxRowsRange(this, false);
        public XlsxCellInteriorRange Interior => new XlsxCellInteriorRange(this);

        public bool FontBold
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue && thisCell.Value.bold;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.bold = value);
            }
        }

        public bool FontItalic
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue && thisCell.Value.italic;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.italic = value);
            }
        }

        public bool FontUnderline
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue && thisCell.Value.underline;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.underline = value);
            }
        }

        public double FontSize
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? sheet.document.fontSize : thisCell.Value.fontSize;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.fontSize = value);
            }
        }

        public string FontName
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? sheet.document.fontName : thisCell.Value.fontName;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.fontName = value);
            }
        }

        public double ColumnWidth
        {
            get
            {
                return new XlsxColumnRange(this, true).ColumnWidth;
            }
            set
            {
                new XlsxColumnRange(this, true).ColumnWidth = value;
            }
        }




        public HorizontalAlignmentValues HorizontalAlignment
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? HorizontalAlignmentValues.General : thisCell.Value.hAlign;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.hAlign = value);
            }
        }

        public VerticalAlignmentValues VerticalAlignment
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? VerticalAlignmentValues.Bottom : thisCell.Value.vAlign;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.vAlign = value);
            }
        }

        public bool WrapText
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? false : thisCell.Value.wrapText;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.wrapText = value);
            }
        }

        public void CenterText()
        {
            HorizontalAlignment = HorizontalAlignmentValues.Center;
            VerticalAlignment = VerticalAlignmentValues.Center;
        }

        public void CenterHorizontal()
        {
            HorizontalAlignment = HorizontalAlignmentValues.Center;
        }

        public object Value
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell?.value;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.value = value);
                //if (value is double || value is float || value is int || value is long)
                //{
                //    NumberFormat = "0.00";
                //}
                //else 
                if (value is string)
                {
                    if ((value as string).IndexOf('\n') >= 0)
                    {
                        WrapText = true;
                    }
                }
                //else if (value is DateTime)
                //{
                //    NumberFormat = "m/d/yyyy";
                //}
                else if (value is XlsxDecoratedText)
                {
                    NumberFormat = "@";
                }
            }
        }

        public string Formula
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell?.formula;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.formula = value);
            }
        }

        public System.Drawing.Color? InteriorColor
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? thisCell.Value.interiorColor : null;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.interiorColor = value);
            }
        }
        public string NumberFormat
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? thisCell.Value.numberFormat : null;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.numberFormat = value);
            }
        }

        public System.Drawing.Color FontColor
        {
            get
            {
                XlsxCell? thisCell = null;
                ProcessRange((ref XlsxCell cell) => thisCell = cell, false);
                return thisCell.HasValue ? thisCell.Value.fontColor : sheet.document.fontColor;
            }
            set
            {
                ProcessRange((ref XlsxCell cell) => cell.fontColor = value);
            }
        }

        public void BorderTable(BorderStyleValues borderWidth = BorderStyleValues.Thin)
        {
            SetBorders(borderWidth,
                XlsxBorder.Top,
                XlsxBorder.Left,
                XlsxBorder.Right,
                XlsxBorder.Bottom,
                XlsxBorder.InsideHorizontal,
                XlsxBorder.InsideVertical);
        }

        public void ClearBorders(params XlsxBorder[] borders)
        {
            SetBorders(BorderStyleValues.None, borders);
        }


        public void SetBorders(BorderStyleValues borderWidth = BorderStyleValues.Thin, params XlsxBorder[] borders)
        {
            for (int i = 0; i < borders.Length; i++)
            {
                switch (borders[i])
                {
                    case XlsxBorder.Top:
                        new XlsxRange(sheet, x, y, width, 1).SetBordersInternal(borderWidth, XlsxBorder.Top);
                        break;
                    case XlsxBorder.Left:
                        new XlsxRange(sheet, x, y, 1, height).SetBordersInternal(borderWidth, XlsxBorder.Left);
                        break;
                    case XlsxBorder.Right:
                        new XlsxRange(sheet, x + width - 1, y, 1, height).SetBordersInternal(borderWidth, XlsxBorder.Right);
                        break;
                    case XlsxBorder.Bottom:
                        new XlsxRange(sheet, x, y + height - 1, width, 1).SetBordersInternal(borderWidth, XlsxBorder.Bottom);
                        break;
                    case XlsxBorder.InsideHorizontal:
                        if (height > 1) new XlsxRange(sheet, x, y, width, height - 1).SetBordersInternal(borderWidth, XlsxBorder.Bottom);
                        break;
                    case XlsxBorder.InsideVertical:
                        if (width > 1) new XlsxRange(sheet, x, y, width - 1, height).SetBordersInternal(borderWidth, XlsxBorder.Right);
                        break;
                }
            }
        }

        private void SetBordersInternal(BorderStyleValues borderWidth = BorderStyleValues.Thin, params XlsxBorder[] borders)
        {
            ProcessRange((ref XlsxCell cell) =>
            {
                for (int i = 0; i < borders.Length; i++)
                {
                    cell.cellBorders.SetBorderWidth(borders[i], borderWidth);
                }
            });

            for (int i = 0; i < borders.Length; i++)
            {
                var border = borders[i];

                for (int cellX = x; cellX < x + width; cellX++)
                {
                    if (border == XlsxBorder.Bottom)
                        sheet.ProcessCell(cellX, y + height, (ref XlsxCell c) => { c.cellBorders.SetBorderWidth(XlsxBorder.Top, borderWidth); }, false);

                    if (y > 1 && border == XlsxBorder.Top)
                        sheet.ProcessCell(cellX, y - 1, (ref XlsxCell c) => { c.cellBorders.SetBorderWidth(XlsxBorder.Bottom, borderWidth); });
                }


                for (int cellY = y; cellY < y + height; cellY++)
                {
                    if (border == XlsxBorder.Right)
                        sheet.ProcessCell(x + width, cellY, (ref XlsxCell c) => { c.cellBorders.SetBorderWidth(XlsxBorder.Left, borderWidth); }, false);

                    if (x > 1 && border == XlsxBorder.Left)
                        sheet.ProcessCell(x - 1, cellY, (ref XlsxCell c) => { c.cellBorders.SetBorderWidth(XlsxBorder.Right, borderWidth); });
                }

            }

        }

        private void ProcessRange(ProcessCellDelegate action, bool create = true)
        {
            for (int cellY = y; cellY < y + height; cellY++)
            {
                for (int cellX = x; cellX < x + width; cellX++)
                {
                    sheet.ProcessCell(cellX, cellY, (ref XlsxCell cell) =>
                    {
                        action(ref cell);
                    }, create);
                }
            }
        }

        public void Merge()
        {
            var thisWidth = width;
            var thisHeight = height;

            sheet.ProcessCell(x, y, (ref XlsxCell cell) =>
            {
                //System.Windows.Forms.MessageBox.Show("Merge(" + x + ", " + y + ", " + cell.colIndex + ", " + cell.rowIndex + ")");

                cell.colSpan = thisWidth;
                cell.rowSpan = thisHeight;
            });


            for (int cellY = y; cellY < y + height; cellY++)
            {
                for (int cellX = x; cellX < x + width; cellX++)
                {
                    if (cellY == y && cellX == x) continue;
                    sheet.ProcessCell(cellX, cellY, (ref XlsxCell cell) =>
                    {
                        cell.mergedToX = cellX;
                        cell.mergedToY = cellY;
                    });
                }
            }
        }

        public void SetValue(object value)
        {
            this.Value = value;
        }

        public void SetFormula(string formula)
        {
            this.Formula = formula;
        }

        public void RowsAutoFit()
        {
            new XlsxRowsRange(this, true).AutoFit();
        }

        public string NoteText
        {
            get
            {
                return sheet.GetComment(XlsxAddress.Address(x, y))?.comment;
            }
            set
            {
                sheet.SetComment(XlsxAddress.Address(x, y), value);
            }
        }

        public void AutoFilter()
        {
            sheet.AutoFilterRange = XlsxAddress.Address(x, y, width, height);
        }
    }

    public class XlsxRowsRange
    {
        XlsxRange range;
        bool entire;

        public XlsxRowsRange(XlsxRange range, bool entire)
        {
            this.range = range;
            this.entire = entire;
        }

        public void AutoFit()
        {
            // TODO
        }

        public double RowHeight
        {
            get
            {
                XlsxRow thisRow = null;
                ProcessRange(row => thisRow = row);
                return thisRow != null && thisRow.rowHeight.HasValue ? thisRow.rowHeight.Value : range.sheet.defaultRowHeight;
            }
            set
            {
                ProcessRange(row => row.rowHeight = value);
            }
        }

        private void ProcessRange(Action<XlsxRow> action)
        {
            for (int rowIndex = range.y; rowIndex < range.y + range.height; rowIndex++)
            {
                var row = range.sheet.GetRow(rowIndex);
                action(row);
            }
        }

        public void Group()
        {
            ProcessRange(row => row.outlineLevel++);
        }
    }

    public class XlsxColumn
    {
        public double? columnWidth;
        public readonly int x;
        public int outlineLevel;

        private XlsxSheet sheet;

        public XlsxColumn(XlsxSheet sheet, int x)
        {
            this.sheet = sheet;
            this.x = x;
            if (sheet.columnsCount < x) sheet.columnsCount = x;
        }
    }

    public class XlsxColumnRange
    {
        XlsxRange range;
        bool entire;

        public XlsxColumnRange(XlsxRange range, bool entire)
        {
            this.range = range;
            this.entire = entire;
        }

        public double ColumnWidth
        {
            get
            {
                XlsxColumn thisColumn = null;
                ProcessRange(col => thisColumn = col);
                return thisColumn != null && thisColumn.columnWidth.HasValue ? thisColumn.columnWidth.Value : 0;
            }
            set
            {
                ProcessRange(col => col.columnWidth = value);
            }
        }

        public bool WrapText
        {
            get
            {
                return new XlsxRange(range.sheet, range.x, 1, range.width, range.sheet.rowsCount).WrapText;
            }
            set
            {
                var newRange = new XlsxRange(range.sheet, range.x, 1, range.width, range.sheet.rowsCount);
                newRange.WrapText = value;
            }
        }

        public void AlignVertical(VerticalAlignmentValues alignment)
        {
            new XlsxRange(range.sheet, range.x, 1, range.width, range.sheet.rowsCount).VerticalAlignment = alignment;
        }

        public void CenterHorizontal()
        {
            new XlsxRange(range.sheet, range.x, 1, range.width, range.sheet.rowsCount).HorizontalAlignment = HorizontalAlignmentValues.Center;
        }

        public void Group()
        {
            ProcessRange(column => column.outlineLevel++);
        }

        private void ProcessRange(Action<XlsxColumn> action)
        {
            for (int x = range.x; x < range.x + range.width; x++)
            {
                var column = range.sheet.GetColumn(x);
                action(column);
            }
        }

        public void AutoFit()
        {
            ProcessRange(column => column.columnWidth = range.sheet.GetMaxColumnWidth(column.x));
        }
    }

    public class XlsxFontRange
    {
        private XlsxRange range;

        public XlsxFontRange(XlsxRange range)
        {
            this.range = range;
        }

        public bool Bold
        {
            get
            {
                return range.FontBold;
            }
            set
            {
                range.FontBold = value;
            }
        }

        public bool Underline
        {
            get
            {
                return range.FontUnderline;
            }
            set
            {
                range.FontUnderline = value;
            }
        }

        public System.Drawing.Color Color
        {
            get
            {
                return range.FontColor;
            }
            set
            {
                range.FontColor = value;
            }
        }

        public double Size
        {
            get
            {
                return range.FontSize;
            }
            set
            {
                range.FontSize = value;
            }
        }

        public string Name
        {
            get
            {
                return range.FontName;
            }
            set
            {
                range.FontName = value;
            }
        }
    }

    public static class XlsxAddress
    {
        public static string Address(int column, int row)
        {
            return string.Format("{0}{1}", Letter(column), row);
        }

        public static string Address(int column, int row, int width, int height)
        {
            return string.Format("{0}:{1}", Address(column, row), Address(column + width - 1, row + height - 1));
        }

        public static string Letter(int col)
        {
            string alphabetName = string.Empty;

            col--;
            do
            {
                alphabetName = (char)('A' + (col % 26)) + alphabetName;
                col = (col / 26) - 1;
            } while (col != -1);

            return alphabetName;
        }

        public static void AddressToColRow(string address, out int col, out int row)
        {
            col = 0;
            address = address.ToUpper();
            StringBuilder numberSB = new StringBuilder();
            for (int i = 0; i < address.Length; i++)
            {
                var ch = address[i];
                if (ch >= 'A') col = (col * 26) + ((int)ch - 64);
                if (char.IsNumber(ch)) numberSB.Append(ch);
            }
            row = int.Parse(numberSB.ToString());
        }
    }

    public class XlsxColor
    {
        public static string ToArgbString(System.Drawing.Color color)
        {
            byte[] values = new byte[] { color.A, color.R, color.G, color.B };
            return BitConverter.ToString(values).Replace("-", string.Empty);
        }

        public static string ToRgbString(System.Drawing.Color color)
        {
            byte[] values = new byte[] { color.R, color.G, color.B };
            return BitConverter.ToString(values).Replace("-", string.Empty);
        }
    }

    public class XlsxDecoratedTextPart
    {
        public string text;
        public System.Drawing.Color? color;
        public bool breakLine;
        //public bool bold;
        //public bool italic;
        //public bool underline;
        //public double? fontSize;


        public override string ToString()
        {
            return $"[\"{text}\"" + (color.HasValue ? ", c:" + XlsxColor.ToArgbString(color.Value) : string.Empty) + (breakLine ? ", break" : string.Empty) + "]";
        }
    }

    public class XlsxDecoratedText
    {
        public List<XlsxDecoratedTextPart> textParts = new List<XlsxDecoratedTextPart>();

        public XlsxDecoratedTextPart AddTextPart(string text)
        {
            var textPart = new XlsxDecoratedTextPart { text = text };
            textParts.Add(textPart);
            return textPart;
        }

        public XlsxDecoratedTextPart AddTextPart(string text, System.Drawing.Color color)
        {
            var textPart = new XlsxDecoratedTextPart { text = text, color = color };
            textParts.Add(textPart);
            return textPart;
        }

        public override string ToString()
        {
            return "{" + string.Join(", ", textParts) + "}";
        }

        internal string GetText()
        {
            return string.Join(string.Empty, textParts.Select(t => t.text));
        }
    }
}
