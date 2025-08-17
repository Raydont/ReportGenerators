using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using InteropExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace MSO.NetTDocs.Excel
{
    public class ExcelReader
    {
        Microsoft.Office.Interop.Excel.Application _excelApp;
        // Workbooks _workBooks;
        Workbook _workBook;
        // Worksheets _workSheets;
        Worksheet _workSheet;
        Range _range;

        public ExcelReader()
        {
            _excelApp = new Microsoft.Office.Interop.Excel.Application(); // new ApplicationClass();
        }

        // открыть документ
        public ExcelReader(string fileName)
        {
            try
            {
                _excelApp = new Microsoft.Office.Interop.Excel.Application(); // new ApplicationClass();
                _workBook = _excelApp.Workbooks.Open(fileName, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _workSheet = (Worksheet)_workBook.Sheets[1];
                _range = _workSheet.get_Range("A1", "A1");
                _excelApp.Visible = true;     // сделать файл Excel видимым
            }
            catch
            {
                MessageBox.Show("Error");                
            }
        }

        // создать документ
        public void NewDocument(string fileName, bool open)
        {
            object missingValue = System.Reflection.Missing.Value;
            _excelApp.DisplayAlerts = false;
            _workBook = _excelApp.Workbooks.Add(missingValue);
            _workSheet = (Worksheet)_workBook.Sheets[1];
            _range = _workSheet.get_Range("A1", "A1");
//            _workBook.CheckCompatibility = false;
            _workBook.SaveAs(fileName, XlFileFormat.xlWorkbookNormal, missingValue, missingValue,
                missingValue, missingValue, XlSaveAsAccessMode.xlExclusive, missingValue, missingValue,
                missingValue, missingValue, missingValue);
            if (open)
            {
                _workBook = _excelApp.Workbooks.Open(fileName, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _workSheet = (Worksheet)_workBook.Sheets[1];
                _range = _workSheet.get_Range("A1", "A1");
                //_excelApp.Visible = true;
            }
            else
            {
                _workBook.Close(true, missingValue, missingValue);
                _excelApp.Quit();
            }
        }

        // сохранение документа
        public void SaveDocument()
        {
             object missingValue = System.Reflection.Missing.Value;
//             _workBook.CheckCompatibility = false;
             _workBook.Close(true, missingValue, missingValue);
             _excelApp.Quit();
             try
             {
                 System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);
                 _excelApp = null;
             }
             catch (Exception ex)
             {
                 _excelApp = null;
                 MessageBox.Show(ex.ToString());
             }
             finally
             {
                 GC.Collect();
             }
        }

        // закрытие файла (без сохранения)
        public void CloseDocument()
        {
            object missingValue = System.Reflection.Missing.Value;
            _workBook.Close(false, missingValue, missingValue);
            _excelApp.Quit();
        }

        // освобождение ресурсов
        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        // защита книги
        public void ProtectWorkbook(string password)
        {
            _workBook.Protect(password, true, false);
        }

        // снятие защиты с книги
        public void UnProtectWorkbook(string password)
        {
            _workBook.Unprotect(password);
        }        

        // изменение текущего(активного) листа в книге без изменения его наименования
        public void ChangeSheet(int sheet)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            ((_Worksheet)_workSheet).Activate();          // сделать текущий лист активным
        }

        // изменение текущего(активного) листа в книге и изменение его наименования (опционально)
        public void ChangeSheet(int sheet, string name)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            ((_Worksheet)_workSheet).Activate();          // сделать текущий лист активным
            if (name != string.Empty) ((_Worksheet)_workSheet).Name = name;
        }

        // добавление листа в текущую книгу и изменение его наименования (опционально)
        public void AddSheet(string name)
        {
            object oldWorksheet = (object)_workSheet;
            _workSheet = (Worksheet)_workBook.Worksheets.Add(Type.Missing, oldWorksheet, Type.Missing, Type.Missing);
            if (name != string.Empty) ((_Worksheet)_workSheet).Name = name;
        }

        // перемещение текущего активного листа в файле Excel
        public void MoveSheet(int sheetBefore)
        {
            _workSheet.Move(Type.Missing, (Worksheet)_workBook.Sheets[sheetBefore]);
        }

        // получение имени листа
        public string GetSheetName(int sheet)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            return ((_Worksheet)_workSheet).Name;
        }

        // защита листа (допускается сортировка)
        public void ProtectSheet(string password)
        {
            _workSheet.Protect(password, true, true, true, false, false, false, false, false, false, false, false, false, false, false, false);
        }

        // снятие защиты листа
        public void UnProtectSheet(string password)
        {
            _workSheet.Unprotect(password);
        }

        // объединение области ячеек
        public void MergeCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Merge(false);
            _range.Select();
            _range.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;
            _range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
        }

        // копирование области ячеек в буфер
        public void CopyCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Copy(Type.Missing);
        }

        // запись области ячеек из буфера в таблицу
        public void PasteCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                false, false);
        }

        // выравнивание текста в ячейке по центру
        public void centerTextCell(string cell)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;
            _range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
        }

        public void CenterText(int row, int col, int height, int width)
        {
            var cell1 = cellNumToAlphabet(row, col);
            var cell2 = cellNumToAlphabet(row + height, col + width);

            _range = _workSheet.get_Range(cell1, cell2);

            _range.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;
            _range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignCenter;
        }

        // добавление рамки к ячейкам
        public void BorderCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.BorderAround(XlLineStyle.xlContinuous, InteropExcel.XlBorderWeight.xlThick, InteropExcel.XlColorIndex.xlColorIndexAutomatic, System.Reflection.Missing.Value);
        }

        public void BorderCells(string cell1, string cell2, InteropExcel.XlBorderWeight weight)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.BorderAround(InteropExcel.XlLineStyle.xlContinuous, weight, InteropExcel.XlColorIndex.xlColorIndexAutomatic, System.Reflection.Missing.Value);
        }
        
        // записать значение в ячейку[row, col]
        public void setValueCell(int row, int col, string value)
        {
            _workSheet.Cells[row,col] = value;            
        }

        // записать значение в ячейку[row, col]
        public void setValueCell(int row, int col, double value)
        {
            _workSheet.Cells[row, col] = value;
        }

        // записать значение в ячейку cell
        public void setValueCell(string cell, string value)
        {
            _range = _workSheet.get_Range(cell,cell);
            _range.Value2 = value;
        }


        // получить значение ячейки[row, col] как string
        public string getValueCell(int row, int col)
        {
            _range = (Range)_workSheet.Cells[row, col];
            if (_range.Value2 == null) return string.Empty;
            return (string)_range.Text;
        }

        // получить значение ячейки cell как string (overload method)
        public string getValueCell(string cell)
        {
            return (string) _workSheet.get_Range(cell, cell).Text;
        }

        // получить ячейку
        public Range getCell(string cell)
        {
            return _workSheet.get_Range(cell, cell);
        }

        // снять защиту диапазона ячеек
        public void UnlockCells(int row, int col, int height, int width)
        {
            var cell1 = cellNumToAlphabet(row, col);
            var cell2 = cellNumToAlphabet(row + height, col + width);
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Locked = false;
            _range.FormulaHidden = false;
        }

        // защита диапазона ячеек
        public void LockCells(int row, int col, int height, int width)
        {
            var cell1 = cellNumToAlphabet(row, col);
            var cell2 = cellNumToAlphabet(row + height, col + width);
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Locked = true;
            _range.FormulaHidden = true;
        }

        // авторазмер ширины столбцов по ширине текста в ячейках
        public void AutoWidth()
        {
            _workSheet.Columns.AutoFit();
        }

        // установка переноса слов в ячейке
        public void WrapText(string cell, bool wrapText)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.WrapText = wrapText;
        }

        // установка переноса слов в ячейке (overload method)
        public void WrapText(int row, int col, bool wrapText)
        {
            _range = (Range)_workSheet.Cells[row, col];
            _range.WrapText = wrapText;            
        }

        public void WrapText(int row, int col, int height, int width,  bool wrapText)
        {
            var cell1 = cellNumToAlphabet(row, col);
            var cell2 = cellNumToAlphabet(row + height, col + width); 

            _range = _workSheet.get_Range(cell1, cell2);

            _range.WrapText = wrapText;
        }

        // жирный шрифт текста в ячейках
        public void Bold(string cell, bool bold)
        {
            _workSheet.get_Range(cell, cell).Font.Bold = bold;
        }
        public void BorderLineCells(int row1, int col1, int row2, int col2, InteropExcel.XlBorderWeight weight, params XlBordersIndex[] borders)
        {
            BorderLineCells(cellNumToAlphabet(row1, col1), cellNumToAlphabet(row2, col2), weight, borders);
        }

        // добавление границ к ячейкам
        public void BorderLineCells(string cell1, string cell2, InteropExcel.XlBorderWeight weight, params XlBordersIndex[] borders)
        {
            _range = _workSheet.get_Range(cell1, cell2);

            // границы

            foreach (var border in borders)
            {
                try
                {
                    _range.Borders[border].LineStyle = XlLineStyle.xlContinuous;
                    _range.Borders[border].Weight = weight;
                }
                catch { }
            }
        }

        public void NumberFormat(int row1, int col1, int row2, int col2, string format)
        {
            NumberFormat(cellNumToAlphabet(row1, col1), cellNumToAlphabet(row2, col2), format);
        }
       
        public void NumberFormat(string cell1, string cell2, string format)
        {
            _range = _workSheet.get_Range(cell1, cell2);

            _range.NumberFormat = format;
        }

        // жирный шрифт текста в ячейках (overload method)
        public void Bold(int row, int col, bool bold)
        {
            string cell = cellNumToAlphabet(row, col);
            Bold(cell, bold);
        }

        // курсивный шрифт текста в ячейках
        public void Italic(string cell, bool italic)
        {
            _workSheet.get_Range(cell, cell).Font.Italic = italic;
        }

        // курсивный шрифт текста в ячейках (overload method)
        public void Italic(int row, int col, bool italic)
        {
            string cell = cellNumToAlphabet(row, col);
            Italic(cell, italic);
        }

        // шрифт с подчеркиванием
        public void Underline(string cell, bool underline)
        {
            _workSheet.get_Range(cell, cell).Font.Underline = underline;
        }

        // шрифт с подчеркиванием (overload method)
        public void Underline(int row, int col, bool underline)
        {
            string cell = cellNumToAlphabet(row, col);
            Underline(cell, underline);
        }

        // цвет ячейки
        public void Color(string cell, int color)
        {
            _workSheet.get_Range(cell, cell).Interior.Color = color;
            if (color == Convert.ToInt32(0xFFFFFF)) _workSheet.get_Range(cell, cell).Interior.ColorIndex = InteropExcel.XlColorIndex.xlColorIndexNone;
        }

        // цвет ячейки (overload method)
        public void Color(int row, int col, int color)
        {
            string cell = cellNumToAlphabet(row, col);
            Color(cell, color);
        }

        // цвет шрифта в ячейке
        public void FontColor(string cell, int color)
        {
            _workSheet.get_Range(cell, cell).Font.Color = color;
           // _workSheet.get_Range(cell, cell).
        }

        // цвет шрифта в ячейке (overload method)
        public void FontColor(int row, int col, int color)
        {
            string cell = cellNumToAlphabet(row, col);
            FontColor(cell, color);
        }

        // закрепление областей на листе Excel
        public void FreezeArea(string cell1, string cell2)
        {
            _workSheet.get_Range(cell1, cell2).Select();
            _excelApp.ActiveWindow.FreezePanes = true;
        }

        // перевод номера столбца в эквивалентное буквенное обозначение Excel
        private string numToAlphabet(int colNumber)
        {
            string alphabetName="";

            colNumber--;
            do {
                alphabetName = (char)('A' + (colNumber % 26)) + alphabetName;
                colNumber=(colNumber/26)-1;
            } while (colNumber!=-1);
            
            return alphabetName;
        }

        // получение номера строки из буквенного обозначения ячейки Excel
        public int getRowFromCell(string cell)
        {
            return _workSheet.get_Range(cell, cell).Row;            
        }

        // получение номера столбца из буквенного обозначения ячейки Excel
        public int getColFromCell(string cell)
        {
            return _workSheet.get_Range(cell, cell).Column;
        }


        // перевод численного обозначения ячейки [row,col] в эквивалетное буквенное обозначение Excel
        public string cellNumToAlphabet(int row, int col)
        {
            return numToAlphabet(col)+row;
        }

        // формула СУММ для указанного столбца(строки) ячеек
        // (записывается в следующую за последней ячейку столбца(строки)
        // down = true|false -  сумма по столбцу|строке  
        // step - шаг смещения вывода суммы по столбцам|строкам
        public string getSUMM(string cell1, string cell2, bool down, int step, int color, bool bold)
        {
            _range = _workSheet.get_Range(cell1,cell2);
            string sumCell = String.Empty;
            if (down)
            {
                int col = _range.Column;
                int row = _range.Row + _range.Count + step;
                sumCell = cellNumToAlphabet(row, col);

//                _workSheet.get_Range(sumCell, sumCell).Formula = "=СУММ(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).FormulaLocal = "=СУММ(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).Font.Bold = bold;
                _workSheet.get_Range(sumCell, sumCell).Font.Color = color;
            }
            else
            {
                int col = _range.Column + _range.Count + step;
                int row = _range.Row;
                sumCell = cellNumToAlphabet(row, col);

//                _workSheet.get_Range(sumCell, sumCell).Formula = "=СУММ(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).FormulaLocal = "=СУММ(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).Font.Bold = bold;
                _workSheet.get_Range(sumCell, sumCell).Font.Color = color;
            }
            return sumCell;
        }

        // произведение значений двух ячеек
        public void getPRODUCT(string targetCell, string prodCell1, string prodCell2)
        {
            _workSheet.get_Range(targetCell, targetCell).Formula = "=" + prodCell1 + "*" + prodCell2;
        }

        // Формула для заданной ячейки
        public void getFormula(string cell, string formula, int color, bool bold)
        {
//            _workSheet.get_Range(cell, cell).Formula = "=" + formula;
            _workSheet.get_Range(cell, cell).FormulaLocal = "=" + formula;
            _workSheet.get_Range(cell, cell).Font.Bold = bold;
            _workSheet.get_Range(cell, cell).Font.Color = color;
        }

        // Группировка по строкам (со скрытием группированных строк)
        public void groupRows(string cell1, string cell2)
        {
            _workSheet.get_Range(cell1, cell2).Rows.Group();
            _workSheet.get_Range(cell1, cell2).EntireRow.Hidden = true;
        }

        // применение стиля
        public void ApplyStyle(int row, int col, dynamic fontName, dynamic fontSize)
        {
            const String STYLE_NAME = "PropertyBorder";
            // Получаем диапазон, содержащий все свойства документа
            Style sty;
            try
            {
                sty = _workBook.Styles[STYLE_NAME];
            }
            catch
            {
                sty = _workBook.Styles.Add(STYLE_NAME, Type.Missing);
            }

            sty.Font.Name = fontName;
            sty.Font.Size = fontSize;

            Range rng = _workSheet.get_Range(cellNumToAlphabet(row, col));
            rng.Style = STYLE_NAME;

            //rng.Columns.AutoFit();
        }

        // применение стиля
        public void ApplyStyle(string alpha, string beta, dynamic fontName, dynamic fontSize)
        {
            const String STYLE_NAME = "PropertyBorder";
            // Получаем диапазон, содержащий все свойства документа
            Style sty;
            try
            {
                sty = _workBook.Styles[STYLE_NAME];
            }
            catch
            {
                sty = _workBook.Styles.Add(STYLE_NAME, Type.Missing);
            }

            sty.Font.Name = fontName;
            sty.Font.Size = fontSize;

            Range rng = _workSheet.get_Range(alpha, beta);
            rng.Style = STYLE_NAME;

            //rng.Columns.AutoFit();
        }

        public void ApplyStyle(dynamic fontName, dynamic fontSize)
        {
            const String STYLE_NAME = "PropertyBorder";
            // Получаем диапазон, содержащий все свойства документа
            Style sty;
            try
            {
                sty = _workBook.Styles[STYLE_NAME];
            }
            catch
            {
                sty = _workBook.Styles.Add(STYLE_NAME, Type.Missing);
            }

            sty.Font.Name = fontName;
            sty.Font.Size = fontSize;
            _range.Style = STYLE_NAME;
            _workSheet.PageSetup.BottomMargin = 18 / 0.268 - 16.8;
            _workSheet.PageSetup.FooterMargin = 22.4;
            _workSheet.PageSetup.RightMargin = 3;
            _workSheet.PageSetup.LeftMargin = 15;
            _workSheet.PageSetup.TopMargin = 18 / 0.268 - 16.8;
            _workSheet.PageSetup.HeaderMargin = 22.4;
        }

        // Изменение ширины столбца
        public void setColumnWidth(string cell, int width)
        {
            _workSheet.get_Range(cell, cell).EntireColumn.ColumnWidth = width;
        }

        public void setColumnWidth(int row, int col, int witdh)
        {
            string cell = cellNumToAlphabet(row, col);
            _workSheet.get_Range(cell, cell).EntireColumn.ColumnWidth = witdh;
        }

        // Задание формата столбца
        public void setColumnFormat(string cell, string format)
        {
            _workSheet.get_Range(cell, cell).EntireColumn.NumberFormat = format;
        }

        public void setColumnFormat(int row, int col, string format)
        {
            string cell = cellNumToAlphabet(row, col);
            _workSheet.get_Range(cell, cell).EntireColumn.NumberFormat = format;
        }

        // выравнивание текста в ячейке по левому/правому краю
        public void alignTextCell(string cell, bool left)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;
            if (left)
                _range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignLeft;
            else
                _range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignRight;
        }

        public void alignTextCell(int row, int col, bool left)
        {
            string cell = cellNumToAlphabet(row, col);
            _range = _workSheet.get_Range(cell, cell);
            _range.VerticalAlignment = InteropExcel.XlVAlign.xlVAlignCenter;
            if (left)
                _range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignLeft;
            else
                _range.HorizontalAlignment = InteropExcel.XlHAlign.xlHAlignRight;
        }

        // Настройки параметров страницы
        public void PageSetup(XlPaperSize paperSize, XlPageOrientation pageOrientation, int leftMargin, int rightMargin, int topMargin, int bottomMargin)
        {
            PageSetup pageSetup = _workSheet.PageSetup;
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

        public dynamic GetSelection()
        {
            return _excelApp.Selection;
        }

        public bool SetPicture(string cell, System.Drawing.Image img, string alternativeText)
        {

            bool result = false;
            try
            {
                // Вставка картинки
                Clipboard.Clear();
                Clipboard.SetImage(img);

                _workSheet.Paste();

                _workSheet.Shapes.Item(_workSheet.Shapes.Count).Select();

                GetSelection().ShapeRange.Left = getCell(cell).Left;
                GetSelection().ShapeRange.Top = getCell(cell).Top;
                GetSelection().ShapeRange.Height = 49F;
                GetSelection().ShapeRange.Width = 85F;
                GetSelection().ShapeRange.AlternativeText = alternativeText;

                GetSelection().ShapeRange.PictureFormat.TransparentBackground = MsoTriState.msoTrue;
                GetSelection().ShapeRange.PictureFormat.TransparencyColor = System.Drawing.Color.White;
                GetSelection().ShapeRange.Fill.Visible = MsoTriState.msoFalse;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка вставки изображения: ");
            }
         result = true;
         return result;
        }
    }
}
