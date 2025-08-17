using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportHelpers
{
    public class DocTableRow
    {
        public List<DocTableCell> cells = new List<DocTableCell>();

        public DocTableCell AddCell(string content, int width = 1, int height = 1, params Action<object>[] applyStyles)
        {
            var cell = new DocTableCell() { content = content, width = width, height = height, applyStyles = applyStyles };
            cells.Add(cell);
            return cell;
        }

        public DocTableCell AddCell(string content, int width = 1, params Action<object>[] applyStyles)
        {
            var cell = new DocTableCell() { content = content, width = width,  applyStyles = applyStyles };
            cells.Add(cell);
            return cell;
        }

        public DocTableCell AddCell(string content, params Action<object>[] applyStyles)
        {
            var cell = new DocTableCell() { content = content, applyStyles = applyStyles };
            cells.Add(cell);
            return cell;
        }
    }

}
