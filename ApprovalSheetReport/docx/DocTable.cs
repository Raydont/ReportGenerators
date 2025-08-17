using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace ReportHelpers
{
    public class DocTable
    {
        private int _width;

        private List<DocTableRow> _rows = new List<DocTableRow>();

        public DocTable(int width)
        {
            _width = width;
        }

        public DocTableRow CreateRow()
        {
            var row = new DocTableRow();
            _rows.Add(row);
            return row;
        }

        public Table BuildTable(WordDoc doc, WdAutoFitBehavior fitBehaviour, Range range = null)
        {
            if (range == null)
            {
                range = doc.Document.Range();
                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
            }
            var table = doc.CreateTable(range, _rows.Count, _width, fitBehaviour);

            for (var i = 0; i < _rows.Count; i++)
            {
                var row = _rows[i];
                int col = 1;
                var rowIndex = i + 1;
                foreach (var cell in row.cells)
                {
                    var docCell = table.Cell(rowIndex, col);
                    docCell.Range.Text = cell.content;
                    foreach (var style in cell.applyStyles)
                    {
                        style(docCell.Range);
                    }

                    col += cell.width;
                }
            }

            for (var i = _rows.Count - 1; i > -1; i--)
            {
                var row = _rows[i];
                var rowIndex = i + 1;

                for (int j = row.cells.Count - 1; j > -1; j--)
                {
                    var cell = row.cells[j];
                    if (cell.width == 1 && cell.height == 1) continue;

                    var col = 1;
                    for (var k = 0; k < row.cells.Count; k++)
                    {
                        var cell1 = row.cells[k];
                        if (cell1 == cell)
                        {
                            break;
                        }
                        else
                        {
                            col += cell1.width;
                        }
                    }


                    var colFrom = col;
                    var rowFrom = rowIndex;
                    var colTo = col + cell.width - 1;
                    var rowTo = rowIndex + cell.height - 1;

                    table.Cell(rowFrom, colFrom).Merge(table.Cell(rowTo, colTo));
                }
            }

            return table;
        }

        public void AddCells(Action<object>[] allCellsFormat, params DocTableCell2[] cells)
        {
            if (allCellsFormat == null) allCellsFormat = new Action<object>[0];

            var countRows = cells.Max(t => t.positionY + t.height - 1);

            var cellFormat = new List<Action<object>>();

            for (var rowIndex = 1; rowIndex <= countRows; rowIndex++)
            {
                var row = CreateRow();

                var colIndex = 1;

                // все ячейки создаваемой строки
                var rowCells = cells.Where(t => t.positionY == rowIndex).OrderBy(t => t.positionX).ToList();

                while (rowCells.Count > 0)
                {
                    var cell = rowCells[0];
                    rowCells.RemoveAt(0);
                    // добавляем пустые ячейки перед вставляемой
                    while (colIndex < cell.positionX)
                    {
                        row.AddCell("", allCellsFormat);
                        colIndex++;
                    }
                    // все стили ячейки
                    cellFormat.Clear();
                    cellFormat.AddRange(allCellsFormat);
                    cellFormat.AddRange(cell.applyStyles);

                    // добавление ячейки
                    row.AddCell(cell.content, cell.width, cell.height, cellFormat.ToArray());
                    colIndex += cell.width;
                }
            }

        }
    }
}
