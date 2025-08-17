using System;
using System.Collections.Generic;
using System.Reflection;
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

        // ������� ��������
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
                _excelApp.Visible = true;     // ������� ���� Excel �������
                
            }
            catch
            {
                MessageBox.Show("������ �������� Excel ���������! ���������� � ���������� ��������������.","Error", MessageBoxButtons.OK , MessageBoxIcon.Error);                
            }
        }

        // ���������� �����
        public void ApplyStyle(int row, int col)
        {
            const String STYLE_NAME = "PropertyBorder";
            // �������� ��������, ���������� ��� �������� ���������
            Style sty;
            try
            {
                sty = _workBook.Styles[STYLE_NAME];
            }
            catch
            {
                sty = _workBook.Styles.Add(STYLE_NAME, Type.Missing);
            }

            sty.Font.Name = "Calibri";
            sty.Font.Size = 11;

            Range rng = _workSheet.get_Range(cellNumToAlphabet(row, col));
            rng.Style = STYLE_NAME;

            //rng.Columns.AutoFit();
        }

        // ���������� �����
        public void ApplyStyle(string alpha, string beta)
        {
            const String STYLE_NAME = "PropertyBorder";
            // �������� ��������, ���������� ��� �������� ���������
            Style sty;
            try
            {
                sty = _workBook.Styles[STYLE_NAME];
            }
            catch
            {
                sty = _workBook.Styles.Add(STYLE_NAME, Type.Missing);
            }

            sty.Font.Name = "Calibri";
            sty.Font.Size = 11;

            Range rng = _workSheet.get_Range(alpha, beta);
            rng.Style = STYLE_NAME;

            //rng.Columns.AutoFit();
        }

        public void ApplyStyle()
        {
            const String STYLE_NAME = "PropertyBorder";
            // �������� ��������, ���������� ��� �������� ���������
            Style sty;
            try
            {
                sty = _workBook.Styles[STYLE_NAME];
            }
            catch
            {
                sty = _workBook.Styles.Add(STYLE_NAME, Type.Missing);
            }

            sty.Font.Name = "Calibri";
            sty.Font.Size = 11;
            _range.Style = STYLE_NAME;
            _workSheet.PageSetup.BottomMargin = 18 / 0.268 - 16.8;
            _workSheet.PageSetup.FooterMargin = 22.4;
            _workSheet.PageSetup.RightMargin = 3;
            _workSheet.PageSetup.LeftMargin = 15;
          //  double topPolePage = 18/0.268;
            _workSheet.PageSetup.TopMargin = 18 / 0.268 - 16.8;
            _workSheet.PageSetup.HeaderMargin = 22.4;


           // _excelApp.Sheets._PrintOut(1, 1, 2, true, Type.Missing, Type.Missing,Type.Missing, Type.Missing);
  //          ((Excel.Worksheet)ThisApplication.Sheets[1]).
  //PrintOut(1, 1, 2, true, Type.Missing, Type.Missing,
  //Type.Missing, Type.Missing);

               // _range = _workSheet.get_Range("A1", "A1");
            
            //Style styForRead = _workBook.Styles[STYLE_NAME];
            //MessageBox.Show("AddIndent "+styForRead.AddIndent.ToString());
          //  MessageBox.Show("Borders " + styForRead.Borders.LineStyle.ToString());
            //MessageBox.Show("FormulaHidden " + styForRead.FormulaHidden.ToString());
            //MessageBox.Show("HorizontalAlignment " + styForRead.HorizontalAlignment.ToString());
            //MessageBox.Show("Orientation " + styForRead.Orientation.ToString());
            //MessageBox.Show("VerticalAlignment " + styForRead.VerticalAlignment.ToString());

            //rng.Columns.AutoFit();
        }
       
        // ������� ��������
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

        // ���������� ���������
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

        // �������� ����� (��� ����������)
        public void CloseDocument()
        {
            object missingValue = System.Reflection.Missing.Value;
            _workBook.Close(false, missingValue, missingValue);
            _excelApp.Quit();
        }

        // ������������ ��������
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

        // ��������� ��������(���������) ����� � ����� ��� ��������� ��� ������������
        public void ChangeSheet(int sheet)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            ((_Worksheet)_workSheet).Activate();          // ������� ������� ���� ��������
        }

        // ��������� ��������(���������) ����� � ����� � ��������� ��� ������������ (�����������)
        public void ChangeSheet(int sheet, string name)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            ((_Worksheet)_workSheet).Activate();          // ������� ������� ���� ��������
            if (name != string.Empty) ((_Worksheet)_workSheet).Name = name;
        }

        // ���������� ����� � ������� ����� � ��������� ��� ������������ (�����������)
        public void AddSheet(string name)
        {
            object oldWorksheet = (object)_workSheet;
            _workSheet = (Worksheet)_workBook.Worksheets.Add(Type.Missing, oldWorksheet, Type.Missing, Type.Missing);
            if (name != string.Empty) ((_Worksheet)_workSheet).Name = name;
        }

        // ����������� �������� ��������� ����� � ����� Excel
        public void MoveSheet(int sheetBefore)
        {
            _workSheet.Move(Type.Missing, (Worksheet)_workBook.Sheets[sheetBefore]);
        }

        // ��������� ����� �����
        public string GetSheetName(int sheet)
        {
            _workSheet = (Worksheet)_workBook.Sheets[sheet];
            return ((_Worksheet)_workSheet).Name;
        }

        // ����������� ������� �����
        public void MergeCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Merge(false);
            _range.Select();
            _range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            _range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        // ����������� ������� ����� � �����
        public void CopyCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.Copy(Type.Missing);
        }

        public void NumberFormat(string cell)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.NumberFormat = "@";
        }

        // ������ ������� ����� �� ������ � �������
        public void PasteCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                false, false);
        }

        // ������������ ������ � ������ �� ������
        public void centerTextCell(string cell)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            _range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }


        // ���������� ������ � �������
        public void BorderLineCells(string cell1, string cell2, params XlBordersIndex[] borders)
        {
            _range = _workSheet.get_Range(cell1, cell2);

            // �������

            foreach (var border in borders)
            {
                _range.Borders[border].LineStyle = XlLineStyle.xlContinuous;
                _range.Borders[border].Weight = XlBorderWeight.xlThin;
            }
        }

        public void centerTextCell(int row, int col)
        {
            centerTextCell(cellNumToAlphabet(row, col));
        }
      
        public void BorderCells(string cell1, string cell2, XlBorderWeight borderWeight)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.BorderAround(XlLineStyle.xlContinuous, borderWeight, XlColorIndex.xlColorIndexAutomatic, System.Reflection.Missing.Value);
        }

        public void BorderLineCells(int row1, int col1, int row2, int col2, params XlBordersIndex[] borders)
        {
            BorderLineCells(cellNumToAlphabet(row1, col1), cellNumToAlphabet(row2, col2), borders);
        }

        public void BorderCells(int row1, int col1, int row2, int col2, XlBorderWeight borderWeight)
        {
            BorderCells(cellNumToAlphabet(row1, col1), cellNumToAlphabet(row2, col2), borderWeight);
        }

        // ���������� ����� � �������
        public void BorderCells(string cell1, string cell2)
        {
            _range = _workSheet.get_Range(cell1, cell2);
            _range.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThick, XlColorIndex.xlColorIndexAutomatic, System.Reflection.Missing.Value);
        }
        
        // �������� �������� � ������[row, col]
        public void setValueCell(int row, int col, string value)
        {
            _workSheet.Cells[row,col] = value;            
        }

        // �������� �������� � ������[row, col]
        public void setValueCell(int row, int col, double value)
        {
            _workSheet.Cells[row, col] = value;
        }

        // �������� �������� � ������ cell
        public void setValueCell(string cell, string value)
        {
            _range = _workSheet.get_Range(cell,cell);
            _range.Value2 = value;
        }


        // �������� �������� ������[row, col] ��� string
        public string getValueCell(int row, int col)
        {
            _range = (Range)_workSheet.Cells[row, col];
            if (_range.Value2 == null) return string.Empty;
            return (string)_range.Text;
        }

        // �������� �������� ������ cell ��� string (overload method)
        public string getValueCell(string cell)
        {
            return (string) _workSheet.get_Range(cell, cell).Text;
        }

        // ���������� ������ �������� �� ������ ������ � �������
        public void AutoWidth()
        {
            _workSheet.Columns.AutoFit();
            _workSheet.Rows.AutoFit();

        }

        // ��������� �������� ���� � ������
        public void WrapText(string cell, bool wrapText)
        {
            _range = _workSheet.get_Range(cell, cell);
            _range.WrapText = wrapText;
        }

        // ��������� �������� ���� � ������ (overload method)
        public void WrapText(int row, int col, bool wrapText)
        {
            _range = (Range)_workSheet.Cells[row, col];
            _range.WrapText = wrapText;            
        }


        // ������ ����� ������ � �������
        public void Bold(string cell, bool bold)
        {
            _workSheet.get_Range(cell, cell).Font.Bold = bold;
        }

        // ������ ����� ������ � ������� (overload method)
        public void Bold(int row, int col, bool bold)
        {
            string cell = cellNumToAlphabet(row, col);
            Bold(cell, bold);
        }

        // ��������� ����� ������ � �������
        public void Italic(string cell, bool italic)
        {
            _workSheet.get_Range(cell, cell).Font.Italic = italic;
        }

        // ��������� ����� ������ � ������� (overload method)
        public void Italic(int row, int col, bool italic)
        {
            string cell = cellNumToAlphabet(row, col);
            Italic(cell, italic);
        }

        // ����� � ��������������
        public void Underline(string cell, bool underline)
        {
            _workSheet.get_Range(cell, cell).Font.Underline = underline;
        }

        // ����� � �������������� (overload method)
        public void Underline(int row, int col, bool underline)
        {
            string cell = cellNumToAlphabet(row, col);
            Underline(cell, underline);
        }

        // ���� ������
        public void Color(string cell, int color)
        {
            _workSheet.get_Range(cell, cell).Interior.Color = color;
            if (color == Convert.ToInt32(0xFFFFFF)) _workSheet.get_Range(cell, cell).Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
        }

        // ���� ������ (overload method)
        public void Color(int row, int col, int color)
        {
            string cell = cellNumToAlphabet(row, col);
            Color(cell, color);
        }

        // ���� ������ � ������
        public void FontColor(string cell, int color)
        {
            _workSheet.get_Range(cell, cell).Font.Color = color;
           // _workSheet.get_Range(cell, cell).
        }

        // ���� ������ � ������ (overload method)
        public void FontColor(int row, int col, int color)
        {
            string cell = cellNumToAlphabet(row, col);
            FontColor(cell, color);
        }

        // ����������� �������� �� ����� Excel
        public void FreezeArea(string cell1, string cell2)
        {
            _workSheet.get_Range(cell1, cell2).Select();
            _excelApp.ActiveWindow.FreezePanes = true;
        }

        // ������� ������ ������� � ������������� ��������� ����������� Excel
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

        // ��������� ������ ������ �� ���������� ����������� ������ Excel
        public int getRowFromCell(string cell)
        {
            return _workSheet.get_Range(cell, cell).Row;            
        }

        // ��������� ������ ������� �� ���������� ����������� ������ Excel
        public int getColFromCell(string cell)
        {
            return _workSheet.get_Range(cell, cell).Column;
        }


        // ������� ���������� ����������� ������ [row,col] � ������������ ��������� ����������� Excel
        public string cellNumToAlphabet(int row, int col)
        {
            return numToAlphabet(col)+row;
        }

        // ���������� ����� � ������

        //public void SetFont(string range, Microsoft.Office.Interop.Excel.Font font)
        //{
        //  //  Range.Font.Name
        //  //  Range = WorkSheet.GetType().InvokeMember("Range", BindingFlags.GetProperty,
        //  //  null, WorkSheet, new object[] { range });

        //    object _Font = _range.GetType().InvokeMember("Font", BindingFlags.GetProperty,
        //    null, _range, null);

        //    _range.GetType().InvokeMember("Name", BindingFlags.SetProperty, null,
        //    _Font, new object[] { font.Name });

        //    _range.GetType().InvokeMember("Size", BindingFlags.SetProperty, null,
        //    _Font, new object[] { font.Size });
        //}

        // ������� ���� ��� ���������� �������(������) �����
        // (������������ � ��������� �� ��������� ������ �������(������)
        // down = true|false -  ����� �� �������|������  
        // step - ��� �������� ������ ����� �� ��������|�������
        public void getSUMM(string cell1, string cell2, bool down, int step, int color, bool bold)
        {
            _range = _workSheet.get_Range(cell1,cell2);
            if (down)
            {
                int col = _range.Column;
                int row = _range.Row + _range.Count + step;
                string sumCell = cellNumToAlphabet(row, col);                

                _workSheet.get_Range(sumCell, sumCell).Formula = "=����(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).Font.Bold = bold;
                _workSheet.get_Range(sumCell, sumCell).Font.Color = color;
            }
            else
            {
                int col = _range.Column + _range.Count + step;
                int row = _range.Row;
                string sumCell = cellNumToAlphabet(row, col);

                _workSheet.get_Range(sumCell, sumCell).Formula = "=����(" + cell1 + ":" + cell2 + ")";
                _workSheet.get_Range(sumCell, sumCell).Font.Bold = bold;
                _workSheet.get_Range(sumCell, sumCell).Font.Color = color;
            }            
        }

        // ������������ �������� ���� �����
        public void getPRODUCT(string targetCell, string prodCell1, string prodCell2)
        {
            _workSheet.get_Range(targetCell, targetCell).Formula = "=" + prodCell1 + "*" + prodCell2;
        }

        // ������� ��� �������� ������
        public void getFormula(string cell, string formula, int color, bool bold)
        {
            _workSheet.get_Range(cell, cell).Formula = "=" + formula;
            _workSheet.get_Range(cell, cell).Font.Bold = bold;
            _workSheet.get_Range(cell, cell).Font.Color = color;
        }
    }
}
