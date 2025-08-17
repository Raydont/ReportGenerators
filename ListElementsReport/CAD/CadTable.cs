using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.Model.Model2D;

namespace Globus.DOCs.Technology.Reports.CAD
{
    public class CadTable
    {
        public Table Table { get; private set; }
        private readonly ParagraphText _text;
        private ParaFormat _paragraphFormat;
        internal CharFormat characterFormat;

        public readonly int countLinesOnFirstPage;
        public readonly int countLinesOnOtherPages;

        public CadTable(Table table, ParagraphText text, int linesOnFirstPage, int linesOnOtherPages)
        {
            countLinesOnFirstPage = linesOnFirstPage;
            countLinesOnOtherPages = linesOnOtherPages;
            Table = table;
            _text = text;
            _paragraphFormat = _text.ParagraphFormat;
            characterFormat = _text.CharacterFormat;
        }

        private readonly List<CadTableRow> _rows = new List<CadTableRow>();

        public CadTableRow CreateRow()
        {
            var row = new CadTableRow(this);
            _rows.Add(row);
            return row;
        }

        public void Apply()
        {
            bool first = true;
            foreach (var row in _rows)
            {
                if (first)
                {
                    first = false;
                    row.Insert(0, Table);
                }
                else
                {
                    AddToTable(row, Table);
                }

            }
            _rows.Reverse();
            _rows.ForEach(MergeCells);
            _rows.Reverse();
        }

        internal void ApplyFormat(ParaFormat.Just just)
        {
            _paragraphFormat.HorJustification = just;
            _paragraphFormat.FitOneLine = true;
            _text.ParagraphFormat = _paragraphFormat;
        }

        internal CharFormat GetCharacterFormat(TextStyle style)
        {
            characterFormat.isBold = false;
            characterFormat.isItalic = false;
            characterFormat.isUnderline = false;
            characterFormat.isStrikeout = false;

            switch (style)
            {
                case TextStyle.Normal:
                    break;
                case TextStyle.Bold:
                    characterFormat.isBold = true;
                    break;
                case TextStyle.Italic:
                    characterFormat.isItalic = true;
                    break;
                case TextStyle.Underline:
                    characterFormat.isUnderline = true;
                    break;
                case TextStyle.BoldItalic:
                    characterFormat.isBold = true;
                    characterFormat.isItalic = true;
                    break;
                case TextStyle.BoldUnderline:
                    characterFormat.isBold = true;
                    characterFormat.isUnderline = true;
                    break;
                case TextStyle.ItalicUnderline:
                    characterFormat.isItalic = true;
                    characterFormat.isUnderline = true;
                    break;
                case TextStyle.BoldItalicUnderline:
                    characterFormat.isBold = true;
                    characterFormat.isItalic = true;
                    characterFormat.isUnderline = true;
                    break;
            }

            return characterFormat;
        }

        public void EmptyRows(int countEmptyRows, int minRowsAfter)
        {
            int lastPageRowsCount = _rows.Count;
            int onPageLeft;
            if ((lastPageRowsCount - countLinesOnFirstPage) > 0)
            {
                lastPageRowsCount -= countLinesOnFirstPage;
                while ((lastPageRowsCount - countLinesOnOtherPages) > 0)
                {
                    lastPageRowsCount -= countLinesOnOtherPages;
                }
                onPageLeft = countLinesOnOtherPages - lastPageRowsCount;
            }
            else
            {
                onPageLeft = countLinesOnFirstPage - lastPageRowsCount;
            }

            if (onPageLeft <= countEmptyRows)
            {
                for (int i = 0; i < onPageLeft; i++)
                {
                    CreateRow();
                }
                return;
            }

            for (int i = 0; i < countEmptyRows; i++)
            {
                CreateRow();
                onPageLeft--;
            }

            if (onPageLeft < minRowsAfter)
            {
                for (int i = 0; i < onPageLeft; i++)
                {
                    CreateRow();
                }
            }
        }



        internal void MergeCells(CadTableRow tableRow)
        {
            uint current = (uint)tableRow.segments.Sum(t => t.length);

            for (int i = tableRow.segments.Count - 1; i >= 0; i--)
            {
                var segment = tableRow.segments[i];
                current -= (uint)segment.length;
                if (segment.length > 1)
                {
                    try
                    {
                        Table.MergeCells(tableRow.StartCell + current, (uint)(tableRow.StartCell + current + segment.length - 1));
                    }
                    catch (Exception)
                    {
                        this.MSG("Invalid cells merge");
                    }

                }
                //Table.SetSelection(current, current);

            }
        }

        internal void AddToTable(CadTableRow tableRow, Table table)
        {
            var lastCell = (uint)Table.CellCount - 1u;
            Table.InsertRows(1, lastCell, Table.InsertProperties.After);
            tableRow.Insert(lastCell + 1, table);
        }

        public int CountEmptyRowsOnPage(int countSpaces = 0)
        {
            for (int i = 0; i < countSpaces; i++)
            {
                _rows.RemoveAt(_rows.Count - 1);
            }

            int lastPageRowsCount = _rows.Count;

            int onPageLeft;
            if ((lastPageRowsCount - countLinesOnFirstPage) > 0)
            {
                lastPageRowsCount -= countLinesOnFirstPage;
                while ((lastPageRowsCount - countLinesOnOtherPages) > 0)
                {
                    lastPageRowsCount -= countLinesOnOtherPages;
                }
                onPageLeft = countLinesOnOtherPages - lastPageRowsCount;
            }
            else
            {
                onPageLeft = countLinesOnFirstPage - lastPageRowsCount;
            }
            return onPageLeft;
        }

        public void NewPage(int countSpaces = 0)
        {
            for (int i = 0; i < countSpaces; i++)
            {
                _rows.RemoveAt(_rows.Count - 1);
            }

            int lastPageRowsCount = _rows.Count;

            int onPageLeft;
            if ((lastPageRowsCount - countLinesOnFirstPage) > 0)
            {
                lastPageRowsCount -= countLinesOnFirstPage;
                while ((lastPageRowsCount - countLinesOnOtherPages) > 0)
                {
                    lastPageRowsCount -= countLinesOnOtherPages;
                }
                onPageLeft = countLinesOnOtherPages - lastPageRowsCount;
            }
            else
            {
                onPageLeft = countLinesOnFirstPage - lastPageRowsCount;
            }

            for (int i = 0; i < onPageLeft; i++)
            {
                CreateRow();
            }
        }
    }
}
