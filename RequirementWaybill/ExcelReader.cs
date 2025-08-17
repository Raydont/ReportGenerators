using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

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
            _excelApp.Application.DisplayAlerts = false;
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
            _workBook = _excelApp.Workbooks.Add(missingValue);
            _workSheet = (Worksheet)_workBook.Sheets[1];
            _range = _workSheet.get_Range("A1", "A1");
            //_workBook.SaveAs(fileName, XlFileFormat.xlWorkbookNormal, missingValue, missingValue,
            //    missingValue, missingValue, XlSaveAsAccessMode.xlExclusive, missingValue, missingValue,
            //    missingValue, missingValue, missingValue);
            if (open)
            {
                _workBook = _excelApp.Workbooks.Open(fileName, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _workSheet = (Worksheet)_workBook.Sheets[1];
                _range = _workSheet.get_Range("A1", "A1");
                // _excelApp.Visible = true;
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

        // закрытие файла (без сохранени€)
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

        // изменение текущего(активного) листа в книге без изменени€ его наименовани€
        public void ChangeSheet(int sheet)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            ((_Worksheet)_workSheet).Activate();          // сделать текущий лист активным
        }

        // изменение текущего(активного) листа в книге и изменение его наименовани€ (опционально)
        public void ChangeSheet(int sheet, string name)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            ((_Worksheet)_workSheet).Activate();          // сделать текущий лист активным
            if (name != string.Empty) ((_Worksheet)_workSheet).Name = name;
        }

        // добавление листа в текущую книгу и изменение его наименовани€ (опционально)
        public void AddSheet(string name)
        {
            object oldWorksheet = (object)_workSheet;
            _workSheet = (Worksheet)_workBook.Worksheets.Add(Type.Missing, oldWorksheet, Type.Missing, Type.Missing);
            if (name != string.Empty) ((_Worksheet)_workSheet).Name = name;
        }

        // ѕеремещение текущего активного листа в файле Excel
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

        // объединение области €чеек
        public void MergeCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Merge(false);
            _range.Select();
            _range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            _range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        // копирование области €чеек в буфер
        public void CopyCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Copy(Type.Missing);
        }

        // запись области €чеек из буфера в таблицу
        public void PasteCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                false, false);
           // _workSheet.get_Range(cell1, cell2).Rows.AutoFit();
            //_range.Rows.AutoFit();
        }

        // выравнивание текста в €чейке по центру
        public void centerTextCell(string cell)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            _range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        // добавление рамки к €чейкам
        public void BorderCells(string cell1)
        {
            _range = _workSheet.get_Range(cell1);
            _range.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, System.Reflection.Missing.Value);
        }

        // изменение высоты строк
        public void AutoFit(int row)
        {

            _workSheet.get_Range(cellNumToAlphabet(row + 8, 12)).RowHeight = 33;
            _workSheet.get_Range(cellNumToAlphabet(row + 14, 6)).RowHeight = 26.25;

        }

        public void AutoFitRow(int row)
        {

            _workSheet.get_Range(cellNumToAlphabet(row, 2), cellNumToAlphabet(row, 12)).RowHeight = 15;

        }

        // изменение высоты строк
        public void SetPrintArea(int row, int col)
        {

            _workSheet.get_Range(cellNumToAlphabet(row + 8, 12)).RowHeight = 33;
            _workSheet.get_Range(cellNumToAlphabet(row + 14, 6)).RowHeight = 26.25;
            _workSheet.PageSetup.PrintArea = cellNumToAlphabet(2, 2) + ":" + cellNumToAlphabet(row, col);

        }

        // изменение шрифта
        public void setFontName(string cell, int size)
        {
            _workSheet.get_Range(cell, cell).Font.Size = size;
            _workSheet.get_Range(cell, cell).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft;
            _workSheet.get_Range(cell, cell).Rows.AutoFit();
        }

       

        public void setFontAmount(string cell, int size)
        {
            _workSheet.get_Range(cell, cell).Font.Size = size;
            _workSheet.get_Range(cell, cell).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;

        }

        public void setFontTotal(string cell, int size)
        {
            _workSheet.get_Range(cell, cell).Font.Size = size;
            _workSheet.get_Range(cell, cell).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight;

        }

        // записать значение в €чейку[row, col]
        public void setValueCell(int row, int col, string value)
        {
            _workSheet.Cells[row,col] = value;            
        }

        // записать значение в €чейку[row, col]
        public void setValueCell(int row, int col, double value)
        {
            _workSheet.Cells[row, col] = value;
        }

        // записать значение в €чейку cell
        public void setValueCell(string cell, string value)
        {
            _range = _workSheet.get_Range(cell,cell);
            _range.Value2 = value;
        }


        // получить значение €чейки[row, col] как string
        public string getValueCell(int row, int col)
        {
            _range = (Range)_workSheet.Cells[row, col];
            if (_range.Value2 == null) return string.Empty;
            return (string)_range.Text;
        }

         //ћен€ем шрифт, можно помен€ть и другие параметры шрифта
        public string setChart(int size, int col)
        {
           // excelapp.
          //  _range = (Range) _workSheet.Cells.BorderAround2( ActiveChart.ChartTitle.Font.Size = 14;
          //  excelapp.ActiveChart.ChartTitle.Font.Color = 255;
            return "s";
        }

        // получить значение €чейки cell как string (overload method)
        public string getValueCell(string cell)
        {
            return (string) _workSheet.get_Range(cell, cell).Text;
        }

        // авторазмер ширины столбцов по ширине текста в €чейках
        public void AutoWidth()
        {
            _workSheet.Columns.AutoFit();
        }

        // установка переноса слов в €чейке
        public void WrapText(string cell, bool wrapText)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.WrapText = wrapText;
        }

        // установка переноса слов в €чейке (overload method)
        public void WrapText(int row, int col, bool wrapText)
        {
            _range = (Range)_workSheet.Cells[row, col];
            _range.WrapText = wrapText;            
        }


        // жирный шрифт текста в €чейках
        public void Bold(string cell, bool bold)
        {
            _workSheet.get_Range(cell, cell).Font.Bold = bold;
        }

        // жирный шрифт текста в €чейках (overload method)
        public void Bold(int row, int col, bool bold)
        {
            string cell = cellNumToAlphabet(row, col);
            Bold(cell, bold);
        }

        // курсивный шрифт текста в €чейках
        public void Italic(string cell, bool italic)
        {
            _workSheet.get_Range(cell, cell).Font.Italic = italic;
        }

        // курсивный шрифт текста в €чейках (overload method)
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

        // цвет €чейки
        public void Color(string cell, int color)
        {
            _workSheet.get_Range(cell, cell).Interior.Color = color;
            if (color == Convert.ToInt32(0xFFFFFF)) _workSheet.get_Range(cell, cell).Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
        }

        // цвет €чейки (overload method)
        public void Color(int row, int col, int color)
        {
            string cell = cellNumToAlphabet(row, col);
            Color(cell, color);
        }

        // цвет шрифта в €чейке
        public void FontColor(string cell, int color)
        {
            _workSheet.get_Range(cell, cell).Font.Color = color;
           // _workSheet.get_Range(cell, cell).
        }

        // цвет шрифта в €чейке (overload method)
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

        // получение номера строки из буквенного обозначени€ €чейки Excel
        public int getRowFromCell(string cell)
        {
            return _workSheet.get_Range(cell, cell).Row;            
        }

        // получение номера столбца из буквенного обозначени€ €чейки Excel
        public int getColFromCell(string cell)
        {
            return _workSheet.get_Range(cell, cell).Column;
        }


        // перевод численного обозначени€ €чейки [row,col] в эквивалетное буквенное обозначение Excel
        public string cellNumToAlphabet(int row, int col)
        {
            return numToAlphabet(col)+row;
        }

        // формула —”ћћ дл€ указанного столбца(строки) €чеек
        // (записываетс€ в следующую за последней €чейку столбца(строки)
        // down = true|false -  сумма по столбцу|строке  
        // step - шаг смещени€ вывода суммы по столбцам|строкам
        public void getSUMM(string cell1, string cell2, bool down, int step, int color, bool bold)
        {
            _range = _workSheet.get_Range(cell1,cell2);
            if (down)
            {
                int col = _range.Column;
                int row = _range.Row + _range.Count + step;
                string sumCell = cellNumToAlphabet(row, col);                

                _workSheet.get_Range(sumCell, sumCell).Formula = "=—”ћћ(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).Font.Bold = bold;
                _workSheet.get_Range(sumCell, sumCell).Font.Color = color;
            }
            else
            {
                int col = _range.Column + _range.Count + step;
                int row = _range.Row;
                string sumCell = cellNumToAlphabet(row, col);

                _workSheet.get_Range(sumCell, sumCell).Formula = "=—”ћћ(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).Font.Bold = bold;
                _workSheet.get_Range(sumCell, sumCell).Font.Color = color;
            }            
        }

        // произведение значений двух €чеек
        public void getPRODUCT(string targetCell, string prodCell1, string prodCell2)
        {
            _workSheet.get_Range(targetCell, targetCell).Formula = "=" + prodCell1 + "*" + prodCell2;
        }

        // ‘ормула дл€ заданной €чейки
        public void getFormula(string cell, string formula, int color, bool bold)
        {
            _workSheet.get_Range(cell, cell).Formula = "=" + formula;
            _workSheet.get_Range(cell, cell).Font.Bold = bold;
            _workSheet.get_Range(cell, cell).Font.Color = color;
        }
    }
}
