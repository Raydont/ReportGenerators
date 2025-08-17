using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LogisticsControlReport
{
    internal class XlsxSpreadsheetGenerator
    {
        public void GenerateFile(string filePath, XlsxDocument documentData)
        {
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    GenerateDocument(documentData, document);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        public byte[] GenerateToMemory(XlsxDocument documentData)
        {
            try
            {
                var start = DateTime.Now;
                using (var ms = new MemoryStream())
                {
                    using (SpreadsheetDocument document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                    {
                        GenerateDocument(documentData, document);
                    }
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
            return null;
        }

        private void GenerateDocument(XlsxDocument documentData, SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();

            workbookPart.Workbook = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15" } };
            workbookPart.Workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbookPart.Workbook.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbookPart.Workbook.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");



            // Добавляем в документ набор стилей
            WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
            wbsp.Stylesheet = GenerateStyleSheet(documentData);
            //wbsp.Stylesheet.Save();

            var workbookViews = new BookViews();
            var view = new WorkbookView();
            workbookViews.Append(view);
            workbookPart.Workbook.Append(workbookViews);

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

            SharedStringTablePart sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            if (sharedStringTablePart.SharedStringTable == null)
            {
                sharedStringTablePart.SharedStringTable = new SharedStringTable();
            }

            var sharedTextValues = new Dictionary<string, int>();


            var calcProps = new CalculationProperties();
            calcProps.FullCalculationOnLoad = true;
            workbookPart.Workbook.AppendChild(calcProps);

            //FileVersion fv = new FileVersion();
            //fv.ApplicationName = "Microsoft Office Excel";

            int sheetId = 1;
            foreach (var worksheetData in documentData.sheets)
            {

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();


                worksheetPart.Worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };

                worksheetPart.Worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                worksheetPart.Worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                worksheetPart.Worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                if (worksheetData.hasComments)
                {
                    GenerateComments(worksheetPart, worksheetData);
                    GenerateVmlDrawingPart1Content(worksheetPart.AddNewPart<VmlDrawingPart>(), worksheetData);
                }

                SheetProperties sheetProperties = new SheetProperties(new OutlineProperties()
                {
                    SummaryBelow = worksheetData.Outline.SummaryRow == XlsxSummaryRow.Below,
                    SummaryRight = worksheetData.Outline.SummaryColumn == XlsxSummaryColumn.Right,
                }
                );

                if (worksheetData.PageSetup.Zoom.HasValue && worksheetData.PageSetup.Zoom.Value)
                {
                    sheetProperties.Append(
                        new PageSetupProperties()
                        {
                            FitToPage = !worksheetData.PageSetup.Zoom.Value
                        }
                    );
                }

                var dimension = new SheetDimension()
                {
                    Reference = XlsxAddress.Address(1, 1, worksheetData.columnsCountWithData, worksheetData.rowsCount)
                };


                var sheetData = new SheetData();

                //Создаем лист в книге
                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)sheetId,
                    Name = worksheetData.Name
                };


                SheetViews sheetViews = new SheetViews();
                SheetView sheetView = new SheetView() { WorkbookViewId = 0 }; // почему всегда 0 хрензнает, но если делаю новые ID, начинает не работать)
                if (worksheetData == documentData.Worksheet)
                {
                    sheetView.TabSelected = true; // активный лист
                }

                if (worksheetData.FreezePanes) // заморозить строку и/или столбец
                {
                    var pane = new Pane()
                    {
                        TopLeftCell = XlsxAddress.Address(worksheetData.SplitColumn + 1, worksheetData.SplitRow + 1),
                        ActivePane = worksheetData.SplitColumn > 0 ? PaneValues.BottomRight : PaneValues.BottomLeft,
                        State = PaneStateValues.Frozen
                    };
                    if (worksheetData.SplitColumn > 0)
                    {
                        pane.HorizontalSplit = worksheetData.SplitColumn;
                    }

                    if (worksheetData.SplitRow > 0)
                    {
                        pane.VerticalSplit = worksheetData.SplitRow;
                    }

                    sheetView.AppendChild(pane);
                }
                sheetViews.Append(sheetView);


                SheetFormatProperties sheetFormatProperties = new SheetFormatProperties();
                sheetFormatProperties.DefaultRowHeight = worksheetData.defaultRowHeight;
                // Аутлайны - группировка колонок
                if (worksheetData.Outline.RowLevels > 0)
                {
                    sheetFormatProperties.OutlineLevelRow = (byte)worksheetData.Outline.RowLevels;
                }
                if (worksheetData.Outline.ColumnLevels > 0)
                {
                    sheetFormatProperties.OutlineLevelColumn = (byte)worksheetData.Outline.ColumnLevels;
                }

                sheetFormatProperties.DefaultColumnWidth = worksheetData.defaultColumnWidth;



                //SheetViews sheetViews1 = new SheetViews();
                //SheetView sheetView1 = new SheetView() { TabSelected = sheetId == 1, WorkbookViewId = (UInt32Value)0U };
                //Selection selection1 = new Selection() { ActiveCell = "A1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" } };
                //sheetView1.Append(selection1);
                //sheetViews1.Append(sheetView1);
                //worksheetPart.Worksheet.AppendChild(sheetViews1);


                var columnsCount = worksheetData.columnsCount;

                // Задаем колонки и их ширину
                Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                bool needToInsertColumns = false;

                if (columnsCount > 0)
                {


                    if (lstColumns == null)
                    {
                        lstColumns = new Columns();
                        needToInsertColumns = true;
                    }

                    for (int i = 0; i < columnsCount; i++)
                    {
                        var colData = worksheetData.GetColumn(i + 1);

                        var col = new Column()
                        {
                            Min = (uint)(i + 1),
                            Max = (uint)(i + 1)
                        };

                        bool hasData = false;

                        if (colData.columnWidth.HasValue)
                        {
                            col.Width = Math.Min(255, colData.columnWidth.Value);
                            //col.BestFit = true;
                            col.CustomWidth = true;
                            hasData = true;
                        }

                        if (colData.outlineLevel > 0)
                        {
                            col.OutlineLevel = (byte)colData.outlineLevel;

                            if (worksheetData.Outline.ColumnLevels <= colData.outlineLevel)
                            {
                                col.Hidden = true;
                            }

                            hasData = true;
                        }

                        if (hasData)
                        {
                            lstColumns.Append(col);
                        }
                    }

                }

                List<XlsxCell> mergedCells = new List<XlsxCell>();


                for (int rowIndex = 0; rowIndex < worksheetData.rowsCount; rowIndex++)
                {
                    Row row = new Row() { RowIndex = (uint)(rowIndex + 1) };

                    var rowInfo = worksheetData.GetRow(rowIndex + 1);
                    if (rowInfo.rowHeight.HasValue)
                    {
                        row.Height = rowInfo.rowHeight.Value;
                        row.CustomHeight = true;
                    }

                    if (rowInfo.outlineLevel > 0)
                    {
                        row.OutlineLevel = (byte)rowInfo.outlineLevel;
                        if (worksheetData.Outline.RowLevels <= rowInfo.outlineLevel)
                        {
                            row.Hidden = true;
                        }
                    }

                    for (int columnIndex = 0; columnIndex < columnsCount; columnIndex++)
                    {
                        worksheetData.ProcessCell(columnIndex + 1, rowIndex + 1, (ref XlsxCell cell) =>
                        {

                            if (cell.colSpan > 1 || cell.rowSpan > 1)
                            {
                                mergedCells.Add(cell);
                            }

                            Cell newCell = new Cell()
                            {
                                CellReference = XlsxAddress.Address(columnIndex + 1, (int)(uint)row.RowIndex),
                                StyleIndex = (uint)cell.styleIndex
                            };



                            if (cell.formula != null)
                            {
                                CellFormula cellFormula1 = new CellFormula();
                                cellFormula1.Text = cell.formula.Trim();
                                if (cellFormula1.Text.StartsWith("=")) cellFormula1.Text = cellFormula1.Text.Substring(1);
                                if (cellFormula1.Text.IndexOf(";") > 0) cellFormula1.Text = cellFormula1.Text.Replace(";", ",");
                                newCell.CellFormula = cellFormula1;
                                newCell.CellFormula.CalculateCell = true;
                            }
                            else if (cell.value == null)
                            {
                                newCell.CellValue = new CellValue();
                            }
                            else if (cell.value is string)
                            {
                                switch (cell.numberFormat)
                                {
                                    case "[h]:mm:ss":
                                        newCell.CellValue = new CellValue(ConvertTime((string)cell.value));
                                        break;
                                    default:
                                        if ((cell.value as string).IndexOf('\n') >= 0)
                                        {
                                            var decoratedText = new XlsxDecoratedText();
                                            decoratedText.AddTextPart(cell.value as string);

                                            newCell.CellValue = new CellValue(InsertSharedString(sharedStringTablePart, sharedTextValues, decoratedText).ToString());
                                            newCell.DataType = CellValues.SharedString;
                                        }
                                        else
                                        {
                                            newCell.CellValue = new CellValue(
                                                InsertSharedString(sharedStringTablePart, sharedTextValues, ReplaceHexadecimalSymbols(cell.value as string)).ToString());
                                            newCell.DataType = CellValues.SharedString;
                                        }
                                        break;
                                }

                                //if (cell.numberFormat == "[h]:mm:ss")
                                //{
                                //    newCell.CellValue = new CellValue(ConvertTime((string)cell.value));
                                //}
                                //else
                                //{

                                //}
                            }
                            else if (cell.value is XlsxDecoratedText)
                            {
                                newCell.CellValue = new CellValue(InsertSharedString(sharedStringTablePart, sharedTextValues, (XlsxDecoratedText)cell.value).ToString());
                                newCell.DataType = CellValues.SharedString;
                            }
                            else if (cell.value is DateTime)
                            {
                                newCell.CellValue = new CellValue(((DateTime)cell.value).ToString("s"));
                                newCell.DataType = CellValues.Date;
                            }
                            else if (cell.value is double || cell.value is float || cell.value is int || cell.value is long)
                            {
                                newCell.CellValue = new CellValue(cell.value.ToString().Replace(",", "."));
                                //newCell.DataType = CellValues.Number;
                            }
                            else
                            {
                                newCell.CellValue = new CellValue(ReplaceHexadecimalSymbols(cell.value + string.Empty));
                                newCell.DataType = CellValues.String;
                            }

                            row.AppendChild(newCell);
                        });
                    }

                    sheetData.Append(row);
                }

                sheets.Append(sheet);
                sheetId++;



                worksheetPart.Worksheet.AppendChild(sheetProperties);
                worksheetPart.Worksheet.AppendChild(dimension);
                worksheetPart.Worksheet.AppendChild(sheetViews);
                worksheetPart.Worksheet.AppendChild(sheetFormatProperties);
                if (needToInsertColumns || lstColumns.ChildElements.Count > 0)
                    worksheetPart.Worksheet.AppendChild(lstColumns);
                worksheetPart.Worksheet.AppendChild(sheetData);


                if (worksheetData.AutoFilterRange != null)
                {
                    AutoFilter autoFilter = new AutoFilter() { Reference = worksheetData.AutoFilterRange };
                    worksheetPart.Worksheet.AppendChild(autoFilter);
                }

                if (mergedCells.Count > 0)
                {
                    MergeCells mergeCells = new MergeCells() { Count = (uint)mergedCells.Count };

                    foreach (var cell in mergedCells)
                    {
                        MergeCell mergeCell = new MergeCell() { Reference = XlsxAddress.Address(cell.colIndex, cell.rowIndex, cell.colSpan, cell.rowSpan) };
                        mergeCells.AppendChild(mergeCell);
                    }

                    worksheetPart.Worksheet.AppendChild(mergeCells);
                }

                PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
                worksheetPart.Worksheet.AppendChild(pageMargins);

                {
                    PageSetup pageSetup1 = new PageSetup()
                    {
                        PaperSize = 9U, // 9 - A4, 8 - A3, 11 - A5
                        Orientation = worksheetData.PageSetup.orientation,
                        HorizontalDpi = 0U,
                        VerticalDpi = 0U
                    };
                    if (worksheetData.PageSetup.FitToPagesTall.HasValue)
                    {
                        pageSetup1.FitToHeight = worksheetData.PageSetup.FitToPagesTall.Value;
                    }

                    if (worksheetData.PageSetup.FitToPagesWide.HasValue)
                    {
                        pageSetup1.FitToWidth = worksheetData.PageSetup.FitToPagesWide.Value;
                    }

                    worksheetPart.Worksheet.AppendChild(pageSetup1);
                }

                if (worksheetData.hasComments)
                {
                    var id = worksheetPart.GetIdOfPart(worksheetPart.GetPartsOfType<VmlDrawingPart>().FirstOrDefault());
                    worksheetPart.Worksheet.Append(new LegacyDrawing() { Id = id });
                }
            }

            //workbookPart.Workbook.Save();
        }

        private void GenerateVmlDrawingPart1Content(VmlDrawingPart vmlDrawingPart1, XlsxSheet sheetData)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);

            StringBuilder sb = new StringBuilder();
            sb.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n" +
                " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n" +
                " xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n" +
                " <o:shapelayout v:ext=\"edit\">\r\n" +
                "  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n" +
                " </o:shapelayout>\r\n" +
                " <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\r\n" +
                "  path=\"m,l,21600r21600,l21600,xe\">\r\n" +
                "  <v:stroke joinstyle=\"miter\"/>\r\n" +
                "  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n" +
                " </v:shapetype>");


            int shapeId = 1025;
            int zindex = 1;
            foreach (var comment in sheetData.GetComments())
            {
                XlsxAddress.AddressToColRow(comment.address, out int col, out int row);

                sb.Append($" <v:shape id=\"_x0000_s{shapeId.ToString("0000")}\" type=\"#_x0000_t202\" style=\'position:absolute;\r\n" +
                    $"  width:108pt;height:59.25pt;z-index:{zindex};\r\n" +
                    "  visibility:hidden\' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\r\n" +
                    "  <v:fill color2=\"#ffffe1\"/>\r\n" +
                    "  <v:shadow color=\"black\" obscured=\"t\"/>\r\n" +
                    "  <v:path o:connecttype=\"none\"/>\r\n" +
                    "  <v:textbox style=\'mso-direction-alt:auto\'>\r\n" +
                    "   <div style=\'text-align:left\'></div>\r\n" +
                    "  </v:textbox>\r\n" +
                    "  <x:ClientData ObjectType=\"Note\">\r\n" +
                    "   <x:MoveWithCells/>\r\n" +
                    "   <x:SizeWithCells/>\r\n" +
                    "   <x:Anchor>\r\n" +
                    $"    {col}, 15, {row - 1}, 10, {col + 2}, 31, {row + 3}, 9</x:Anchor>\r\n" +
                    "   <x:AutoFill>False</x:AutoFill>\r\n" +
                    $"   <x:Row>{row - 1}</x:Row>\r\n" +
                    $"   <x:Column>{col - 1}</x:Column>\r\n" +
                    "  </x:ClientData>\r\n" +
                    " </v:shape>"); ;
                shapeId++;
                zindex++;
            }
            sb.Append("</xml>");

            writer.WriteRaw(sb.ToString());
            writer.Flush();
            writer.Close();
        }

        private void GenerateComments(WorksheetPart worksheetPart, XlsxSheet sheetData)
        {
            WorksheetCommentsPart worksheetCommentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();

            Comments comments = new Comments();

            Authors authors = new Authors();
            Author author1 = new Author();
            author1.Text = "Генератор отчета";
            authors.Append(author1);

            CommentList commentList = new CommentList();

            foreach (var commentInfo in sheetData.GetComments())
            {
                Comment comment = new Comment() { Reference = commentInfo.address, AuthorId = 0U, ShapeId = 0U };

                CommentText commentText = new CommentText();

                //Run run1 = new Run();

                //RunProperties runProperties1 = new RunProperties();
                //runProperties1.Append(new Bold());
                //runProperties1.Append(new FontSize() { Val = 9D });
                //runProperties1.Append(new Color() { Indexed = 81U });
                //runProperties1.Append(new RunFont() { Val = "Tahoma" });
                //runProperties1.Append(new RunPropertyCharSet() { Val = 1 });
                //Text text1 = new Text();
                //text1.Text = "Генератор отчета";
                //run1.Append(runProperties1);
                //run1.Append(text1);

                Run run2 = new Run();

                RunProperties runProperties2 = new RunProperties();

                runProperties2.Append(new FontSize() { Val = 9D });
                runProperties2.Append(new Color() { Indexed = 81U });
                runProperties2.Append(new RunFont() { Val = "Tahoma" });
                runProperties2.Append(new RunPropertyCharSet() { Val = 1 });
                Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text2.Text = /*"\n" +*/ commentInfo.comment;

                run2.Append(runProperties2);
                run2.Append(text2);

                //commentText.Append(run1);
                commentText.Append(run2);

                comment.Append(commentText);

                commentList.Append(comment);
            }

            comments.Append(authors);
            comments.Append(commentList);

            worksheetCommentsPart.Comments = comments;
        }

        private string ConvertTime(string strVal)
        {
            var parts = strVal.Split(':');
            if (parts.Length > 1)
            {
                var mul = 1d;
                var value = 0d;
                for (int i = parts.Length - 1; i > -1; i--)
                {
                    value += int.Parse(parts[i]) * mul / 86400d;
                    mul *= 60;
                }
                return value.ToString().Replace(",", ".");
            }
            else
            {
                return strVal;
            }
        }

        int InsertSharedString(SharedStringTablePart sharedStringTablePart, Dictionary<string, int> sharedTextValues, string text)
        {
            if (sharedTextValues.TryGetValue(text, out int index))
            {
                return index;
            }

            sharedTextValues.Add(text, sharedTextValues.Count);
            sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

            return sharedTextValues.Count - 1;
        }

        int InsertSharedString(SharedStringTablePart sharedStringTablePart, Dictionary<string, int> sharedTextValues, XlsxDecoratedText text)
        {
            var textString = text.ToString();

            if (sharedTextValues.TryGetValue(textString, out int index))
            {
                return index;
            }


            sharedTextValues.Add(textString, sharedTextValues.Count);

            var sharedStringItemElement = new SharedStringItem();


            foreach (var textPart in text.textParts)
            {
                string[] texts;
                if (textPart.text.IndexOf('\n') >= 0)
                {

                    var text1 = textPart.text.Replace("\r\n", "\n").Replace("\n", "\r\n");

                    texts = text1.Split(new char[] { '\r' }, StringSplitOptions.None);
                }
                else
                {
                    texts = new string[] { textPart.text };
                }

                foreach (var t in texts)
                {
                    var run = new Run();
                    if (textPart.color.HasValue)
                    {
                        var props = new RunProperties();
                        props.AppendChild(new Color() { Rgb = XlsxColor.ToArgbString(textPart.color.Value) });
                        run.AppendChild(props);
                    }


                    var textE = new Text(ReplaceHexadecimalSymbols(t));
                    //if (t.Trim() != t)
                    {
                        textE.Space = SpaceProcessingModeValues.Preserve;
                    }
                    run.AppendChild(textE);
                    sharedStringItemElement.AppendChild(run);
                }

            }
            sharedStringTablePart.SharedStringTable.AppendChild(sharedStringItemElement);

            return sharedTextValues.Count - 1;
        }


        //Важный метод, при вставки текстовых значений надо использовать.
        //Метод убирает из строки запрещенные спец символы.
        //Если не использовать, то при наличии в строке таких символов, вылетит ошибка.
        static Regex r = new Regex("[\x00-\x08\x0B\x0C\x0E-\x1F\x26]", RegexOptions.Compiled);
        static string ReplaceHexadecimalSymbols(string txt)
        {
            return r.Replace(txt, string.Empty);
        }

        class FontInfo
        {
            public string name;
            public double size;
            public System.Drawing.Color color;
            public bool bold;
            public bool italic;
            public bool underline;
        }

        class BorderInfo
        {
            public BorderStyleValues top = BorderStyleValues.None;
            public BorderStyleValues bottom = BorderStyleValues.None;
            public BorderStyleValues left = BorderStyleValues.None;
            public BorderStyleValues right = BorderStyleValues.None;
        }

        class CellAllignmentInfo
        {
            public VerticalAlignmentValues vAlign;
            public HorizontalAlignmentValues hAlign;
        }

        class NumberFormatInfo
        {
            public string formatCode;
        }

        class FillInfo
        {
            public System.Drawing.Color? interiorColor;
            internal string pattern;

            internal string GetInteriorColor()
            {
                byte[] values = new byte[] { interiorColor.Value.A, interiorColor.Value.R, interiorColor.Value.G, interiorColor.Value.B };
                return BitConverter.ToString(values).Replace("-", string.Empty);
            }
        }

        class CellStyleInfo
        {
            public int fontIndex;
            public int borderIndex;
            public int alignmentIndex;
            public int numberFormatIndex;
            public int fillIndex;
            public bool wrapText;
        }

        List<FontInfo> fonts = new List<FontInfo>();
        List<BorderInfo> borders = new List<BorderInfo>();
        List<CellAllignmentInfo> alignments = new List<CellAllignmentInfo>();
        List<CellStyleInfo> styles = new List<CellStyleInfo>();
        List<NumberFormatInfo> numberFormats = new List<NumberFormatInfo>();
        List<FillInfo> fills = new List<FillInfo>();

        T FirstOrDefault<T>(List<T> list, Func<T, bool> action, int startIndex = 0)
        {
            for (int i = startIndex; i < list.Count; i++)
            {
                if (action(list[i]))
                {
                    return list[i];
                }
            }
            return default;
        }

        //Метод генерирует стили для ячеек (за основу взят код, найденный где-то в интернете)
        Stylesheet GenerateStyleSheet(XlsxDocument documentData)
        {
            borders.Add(new BorderInfo());

            styles.Add(new CellStyleInfo { alignmentIndex = -1, borderIndex = 0, fontIndex = 0, wrapText = false });

            fonts.Add(new FontInfo { color = documentData.fontColor, name = documentData.fontName, size = documentData.fontSize });

            numberFormats.Add(new NumberFormatInfo { formatCode = string.Empty });

            fills.Add(new FillInfo { interiorColor = null });
            fills.Add(new FillInfo { interiorColor = null, pattern = "grey125" });

            for (int i = 0; i < documentData.sheets.Count; i++)
            {
                var sheet = documentData.sheets[i];
                var columnsCount = sheet.columnsCount;
                for (int row = 1; row <= sheet.rowsCount; row++)
                {
                    for (int col = 1; col <= columnsCount; col++)
                    {
                        sheet.ProcessCell(col, row, (ref XlsxCell c) =>
                        {
                            var cell = c;

                            var font = FirstOrDefault(fonts, t =>
                                 t.name == cell.fontName &&
                                 Math.Abs(t.size - cell.fontSize) < 0.001 &&
                                 t.color == cell.fontColor &&
                                 t.bold == cell.bold &&
                                 t.italic == cell.italic &&
                                 t.underline == cell.underline
                                );

                            if (font == null)
                            {
                                font = new FontInfo
                                {
                                    size = cell.fontSize,
                                    name = cell.fontName,
                                    color = cell.fontColor,
                                    bold = cell.bold,
                                    italic = cell.italic,
                                    underline = cell.underline
                                };
                                fonts.Add(font);
                            }

                            var bordersInfo = FirstOrDefault(borders, t => t.top == cell.cellBorders.GetBorderWidth(XlsxBorder.Top) &&
                                 t.bottom == cell.cellBorders.GetBorderWidth(XlsxBorder.Bottom) &&
                                 t.left == cell.cellBorders.GetBorderWidth(XlsxBorder.Left) &&
                                 t.right == cell.cellBorders.GetBorderWidth(XlsxBorder.Right));

                            if (bordersInfo == null)
                            {
                                bordersInfo = new BorderInfo
                                {
                                    top = cell.cellBorders.GetBorderWidth(XlsxBorder.Top),
                                    bottom = cell.cellBorders.GetBorderWidth(XlsxBorder.Bottom),
                                    left = cell.cellBorders.GetBorderWidth(XlsxBorder.Left),
                                    right = cell.cellBorders.GetBorderWidth(XlsxBorder.Right)
                                };
                                borders.Add(bordersInfo);
                            }

                            var alignment = FirstOrDefault(alignments, t => t.hAlign == cell.hAlign && t.vAlign == cell.vAlign);
                            if (alignment == null)
                            {
                                alignment = new CellAllignmentInfo { hAlign = cell.hAlign, vAlign = cell.vAlign };
                                alignments.Add(alignment);
                            }

                            var numFormat = FirstOrDefault(numberFormats, t => t.formatCode == (cell.numberFormat ?? string.Empty));
                            if (numFormat == null)
                            {
                                numFormat = new NumberFormatInfo { formatCode = cell.numberFormat };
                                numberFormats.Add(numFormat);
                            }

                            var fill = FirstOrDefault(fills, t => t.interiorColor == cell.interiorColor && t.pattern == null);
                            if (fill == null)
                            {
                                fill = new FillInfo { interiorColor = cell.interiorColor };
                                fills.Add(fill);
                            }

                            var fontIndex = fonts.IndexOf(font);
                            var borderIndex = borders.IndexOf(bordersInfo);
                            var alignmentIndex = alignments.IndexOf(alignment);
                            var numberFormatIndex = numberFormats.IndexOf(numFormat);
                            var fillIndex = fills.IndexOf(fill);

                            var style = FirstOrDefault(styles, t =>
                                 t.alignmentIndex == alignmentIndex &&
                                 t.borderIndex == borderIndex &&
                                 t.fontIndex == fontIndex &&
                                 t.numberFormatIndex == numberFormatIndex &&
                                 t.fillIndex == fillIndex &&
                                 t.wrapText == cell.wrapText, 1);

                            if (style == null)
                            {
                                style = new CellStyleInfo()
                                {
                                    alignmentIndex = alignmentIndex,
                                    borderIndex = borderIndex,
                                    fontIndex = fontIndex,
                                    numberFormatIndex = numberFormatIndex,
                                    wrapText = cell.wrapText,
                                    fillIndex = fillIndex
                                };
                                styles.Add(style);
                            }

                            c.styleIndex = styles.IndexOf(style);
                        });
                    }

                    //if (Globus.DOCs.Technology.Reports.ServerSide.Dialogs.ProgressDialog.instance != null)
                    //{
                    //    Globus.DOCs.Technology.Reports.ServerSide.Dialogs.ProgressDialog.instance.SetInfo("GenerateStyleSheet(row: " + row
                    //        + ", fonts:" + fonts.Count
                    //        + ", borders:" + borders.Count
                    //        + ", alignments:" + alignments.Count
                    //        + ", numberFormats:" + numberFormats.Count
                    //        + ", fills:" + fills.Count
                    //        + ", styles:" + styles.Count
                    //        + ", columnsCount:" + columnsCount + ", " + sheet.columnsCount
                    //        + ")");
                    //}
                }
            }




            var stylesheet = new Stylesheet();

            // шрифты
            var fontsElement = new Fonts();
            fontsElement.Count = (uint)fonts.Count;
            foreach (var fontInfo in fonts)
            {
                var font = new Font();

                if (fontInfo.bold) font.AppendChild(new Bold());
                if (fontInfo.italic) font.AppendChild(new Italic());
                if (fontInfo.underline) font.AppendChild(new Underline());

                font.AppendChild(new FontSize() { Val = fontInfo.size });
                font.AppendChild(new Color() { Rgb = new HexBinaryValue() { Value = XlsxColor.ToArgbString(fontInfo.color) } });
                font.AppendChild(new FontName() { Val = fontInfo.name });

                fontsElement.AppendChild(font);
            }
            stylesheet.AppendChild(fontsElement);

            // форматы чисел
            stylesheet.NumberingFormats = new NumberingFormats();
            stylesheet.NumberingFormats.Count = (uint)numberFormats.Count;
            int nfid = 0;
            foreach (var numberFormat in numberFormats)
            {
                var nFormat = new NumberingFormat();
                nFormat.FormatCode = numberFormat.formatCode;
                nFormat.NumberFormatId = (uint)nfid;
                stylesheet.NumberingFormats.AppendChild(nFormat);
                nfid++;
            }


            // заливки
            var fillsElement = new Fills();
            fillsElement.Count = (uint)fills.Count;

            foreach (var fill in fills)
            {
                var fillElement = new Fill();
                var fillPattern = new PatternFill();

                if (fill.pattern != null)
                {
                    fillPattern.PatternType = PatternValues.Gray125;
                }
                else
                {
                    if (fill.interiorColor.HasValue)
                    {
                        fillPattern.ForegroundColor = new ForegroundColor() { Rgb = fill.GetInteriorColor() };
                        fillPattern.PatternType = PatternValues.Solid;
                    }
                    else
                    {
                        fillPattern.ForegroundColor = new ForegroundColor() { Auto = true };
                        fillPattern.PatternType = PatternValues.None;
                    }
                }
                fillElement.AppendChild(fillPattern);

                fillsElement.AppendChild(fillElement);
            }
            stylesheet.AppendChild(fillsElement);

            // границы
            var bordersElement = new Borders();
            bordersElement.Count = (uint)borders.Count;

            foreach (var borderInfo in borders)
            {
                var border = new Border();
                if (borderInfo.left == BorderStyleValues.None)
                {
                    border.AppendChild(new LeftBorder());
                }
                else
                {
                    border.AppendChild(new LeftBorder(new Color() { Auto = true }) { Style = borderInfo.left });
                }

                if (borderInfo.right == BorderStyleValues.None)
                {
                    border.AppendChild(new RightBorder());
                }
                else
                {
                    border.AppendChild(new RightBorder(new Color() { Auto = true }) { Style = borderInfo.right });
                }

                if (borderInfo.top == BorderStyleValues.None)
                {
                    border.AppendChild(new TopBorder());
                }
                else
                {
                    border.AppendChild(new TopBorder(new Color() { Auto = true }) { Style = borderInfo.top });
                }
                if (borderInfo.bottom == BorderStyleValues.None)
                {
                    border.AppendChild(new BottomBorder());
                }
                else
                {
                    border.AppendChild(new BottomBorder(new Color() { Auto = true }) { Style = borderInfo.bottom });
                }

                border.AppendChild(new DiagonalBorder());

                bordersElement.AppendChild(border);
            }
            stylesheet.AppendChild(bordersElement);

            // форматы стилей
            var styleElement = new CellStyleFormats();
            styleElement.Count = 1;
            {
                var style = styles.FirstOrDefault(t => t.alignmentIndex == -1);
                var cellFormat = new CellFormat();
                cellFormat.FontId = (uint)style.fontIndex;
                cellFormat.FillId = 0;
                cellFormat.BorderId = (uint)style.borderIndex;
                styleElement.AppendChild(cellFormat);

                cellFormat.AppendChild(new Alignment()
                {
                    WrapText = true
                });
            }
            stylesheet.AppendChild(styleElement);

            // форматы ячеек
            var cellFormatsElement = new CellFormats();
            cellFormatsElement.Count = (uint)styles.Count;
            foreach (var style in styles)
            {
                var cellFormat = new CellFormat();
                cellFormat.FormatId = 0;

                cellFormat.FontId = (uint)style.fontIndex;
                cellFormat.FillId = (uint)style.fillIndex;
                cellFormat.BorderId = (uint)style.borderIndex;
                cellFormat.NumberFormatId = (uint)style.numberFormatIndex;


                if (style.alignmentIndex >= 0)
                {
                    var alignment = alignments[style.alignmentIndex];

                    cellFormat.AppendChild(new Alignment()
                    {
                        Horizontal = alignment.hAlign,
                        Vertical = alignment.vAlign,
                        WrapText = style.wrapText
                    });

                    cellFormat.ApplyAlignment = alignment.hAlign != HorizontalAlignmentValues.General || alignment.vAlign != VerticalAlignmentValues.Bottom;
                    cellFormat.ApplyBorder = true;
                    cellFormat.ApplyFont = true;
                    cellFormat.ApplyNumberFormat = true;
                    cellFormat.ApplyFill = true;
                }


                cellFormatsElement.AppendChild(cellFormat);

            }
            stylesheet.AppendChild(cellFormatsElement);

            // стили
            var cellStylesElement = new CellStyles(new CellStyle { Name = "Обычный", FormatId = 0, BuiltinId = 0 });
            cellStylesElement.Count = 1;
            stylesheet.AppendChild(cellStylesElement);

            return stylesheet;
        }

    }
}
