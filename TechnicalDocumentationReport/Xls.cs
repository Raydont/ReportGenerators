using System;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ReportHelpers
{
    public class Xls : IDisposable
    {
        private Application _application;
        public Application Application
        {
            get { return _application ?? (_application = new Application { Visible = _visible }); }
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

        private Worksheet _worksheet;
        public Worksheet Worksheet
        {
            get { return _worksheet ?? (_worksheet = WorkBook.Sheets[1]); }
        }


        private Workbook _workBook;
        public Workbook WorkBook
        {
            get { return _workBook ?? (_workBook = Application.Workbooks.Add(Type.Missing)); }
        }

        public void Open(string workBookPath, bool saveLast = true)
        {
            if (_workBook != null && !string.IsNullOrWhiteSpace(_workBook.Path))
            {
                _workBook.Close(saveLast);
                _workBook = null;
            }
            _worksheet = null;
            _workBook = Application.Workbooks.Open(workBookPath);
        }

        public void Close(bool saveChanges = false)
        {
            if (_workBook != null)
            {
                _workBook.Close(saveChanges);
                _worksheet = null;
                _workBook = null;
            }
        }

        public void Save()
        {
            if (_workBook != null) _workBook.Save();
        }

        public void SaveAs(string workBookPath)
        {
            if (_workBook != null) _workBook.SaveAs(workBookPath);
        }

        public void Dispose()
        {
            Quit();
        }

        public void Quit(bool saveChanges = false)
        {
            if (_application == null) return;

            if (_workBook != null)
            {
                _workBook.Close(saveChanges);
                _workBook = null;
            }

            _application.Quit();
            _application = null;
        }

        public void SelectWorksheet(int index)
        {
            if (index < 1 || index > WorkBook.Sheets.Count)
            {
                throw new InvalidOperationException("Индкс за педелами коллекции!");
            }

            _worksheet = WorkBook.Sheets[index];
        }

        public void SelectWorksheet(string name)
        {
            foreach (Worksheet sheet in WorkBook.Sheets)
            {
                if (sheet.Name == name)
                {
                    _worksheet = sheet;
                    return;
                }
            }

            throw new InvalidOperationException("Лист \"" + name + "\" не найден!");
        }

        public void AddWorksheet(string name, bool select)
        {
            var newWorksheet = (Worksheet)WorkBook.Sheets.Add(Type.Missing, Worksheet ?? Type.Missing, Type.Missing, Type.Missing);
            newWorksheet.Name = name;
            if (select)
            {
                _worksheet = newWorksheet;
            }
        }

        public static string Address(int column, int row)
        {
            return string.Format("{0}{1}", Letter(column), row);
        }

        public static string Address(int column, int row, int width, int height)
        {
            return string.Format("{0}:{1}", Address(column, row), Address(column + width - 1, row + height - 1));
        }

        public Range this[int column, int row]
        {
            get
            {
                return Worksheet.Range[Address(column, row)];
            }
        }

        public Range this[int column, int row, int width, int height]
        {
            get
            {
                return Worksheet.Range[Address(column, row, width, height)];
            }
        }

        public Range this[string address]
        {
            get
            {
                return Worksheet.Range[address];
            }
        }

        public static string Letter(int col)
        {
            string alphabetName = "";

            col--;
            do
            {
                alphabetName = (char)('A' + (col % 26)) + alphabetName;
                col = (col / 26) - 1;
            } while (col != -1);

            return alphabetName;
        }

        public void AutoWidth()
        {
            Worksheet.Columns.AutoFit();
        }
    }

    public static class XlsExtender
    {
        public static void CenterText(this Range range)
        {
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public static void Border(this Range range, XlLineStyle style = XlLineStyle.xlContinuous, XlBorderWeight weight = XlBorderWeight.xlMedium)
        {
            range.BorderAround(style, weight);
        }

        static readonly XlBordersIndex[] _tableBorders =
        {
            XlBordersIndex.xlEdgeBottom, 
            XlBordersIndex.xlEdgeTop, 
            XlBordersIndex.xlEdgeLeft, 
            XlBordersIndex.xlEdgeRight,
            XlBordersIndex.xlInsideHorizontal, 
            XlBordersIndex.xlInsideVertical
        };

        public static void BorderTable(this Range range, XlLineStyle style = XlLineStyle.xlContinuous, XlBorderWeight weight = XlBorderWeight.xlMedium)
        {
            SetBorders(range, style, weight, _tableBorders);
        }

        public static void SetBorders(this Range range, XlLineStyle style = XlLineStyle.xlContinuous, XlBorderWeight weight = XlBorderWeight.xlMedium,
            params XlBordersIndex[] borders)
        {
            foreach (var xlBordersIndex in borders)
            {
                var border = range.Borders[xlBordersIndex];
                try
                {
                    border.LineStyle = style;
                }
                // ReSharper disable EmptyGeneralCatchClause
                catch
                {
                }
                // ReSharper restore EmptyGeneralCatchClause                
                border.Weight = weight;
            }
        }

        public static void SetValue(this Range range, object value)
        {
            range.Value[XlRangeValueDataType.xlRangeValueDefault] = value;
        }


    }

}
