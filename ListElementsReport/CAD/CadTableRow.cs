using System.Collections.Generic;
using TFlex.Model.Model2D;

namespace Globus.DOCs.Technology.Reports.CAD
{
    public class CadTableRow
    {

        private readonly CadTable _table;
        internal CadTableRow(CadTable table)
        {
            _table = table;
        }

        internal class RowSegment
        {
            public string text;
            public int length;
            public CadTableCellJust just = CadTableCellJust.Left;
            public TextStyle style = TextStyle.Normal;

            public override string ToString()
            {
                return text + "\r\n" + "Length: " + length + "\r\n" + "Just: " + just + "\r\n" + "Style: " + style;
            }
        }

        internal readonly List<RowSegment> segments = new List<RowSegment>();

        internal uint StartCell { get; private set; }


        public void AddText(string text, CadTableCellJust cellJust = CadTableCellJust.Left, TextStyle textStyle = TextStyle.Normal)
        {
            segments.Add(new RowSegment { text = text, length = 1, just = cellJust, style = textStyle });
        }

        public void AddText(string text, int countCells, CadTableCellJust cellJust = CadTableCellJust.Left, TextStyle textStyle = TextStyle.Normal)
        {
            segments.Add(new RowSegment { text = text, length = countCells, just = cellJust, style = textStyle});
        }

        internal void Insert(uint startCell, Table table)
        {
            StartCell = startCell;
            uint current = 0;
            foreach (var segment in segments)
            {
                table.InsertText(startCell + current, 0, segment.text, _table.GetCharacterFormat(segment.style));
                //MessageBox.Show("index: " + (startCell + current) + "\r\n" + segment);

                var just = ParaFormat.Just.Left;
                switch (segment.just)
                {
                    case CadTableCellJust.Left:
                        just = ParaFormat.Just.Left;
                        break;
                    case CadTableCellJust.Center:
                        just = ParaFormat.Just.Center;
                        break;
                    case CadTableCellJust.Right:
                        just = ParaFormat.Just.Right;
                        break;
                }
                _table.ApplyFormat(just);

                current += (uint)segment.length;
            }
        }


    }
}
