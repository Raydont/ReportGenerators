using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Model;
using TFlex.Model.Model2D;

namespace BuyProductsReport
{
    public static class Extenders
    {
        public static string GetStringOrEmpty(this Hashtable hashtable, object key)
        {
            if (hashtable.ContainsKey(key))
            {
                var val = hashtable[key];
                if (val != null)
                {
                    var str = val.ToString().Trim();
                    return str == "0" ? "" : str;
                }
            }

            return "";
        }

        public static List<Fragment> GetFragments(this Page pPage)
        {
            if (pPage == null)
            {
                return null;
            }

            ICollection<Fragment> fragments = pPage.Document.GetFragments();
            if (fragments == null)
            {
                return null;
            }

            return fragments.Where(fragment => fragment.Page == pPage).ToList();
        }

        public static Page InsertNewPage(this Document document)
        {
            Page result = null;
            try
            {
                document.BeginChanges("Вставка новой страницы");
                result = new Page(document);
            }
            catch (Exception ex)
            {
                string text = "InsertNewPage()\n" + ex.Message;
                MessageBox.Show(text, "Ошибка создания новой страницы", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
            finally
            {
                document.EndChanges();
            }
            return result;
        }

        public static int CountPages(this Document document)
        {
            var pages = document.GetPages();
            return pages.Count;
        }

        public static void MSG(this object obj, string msg)
        {
            MessageBox.Show(msg);
        }


        public static ParagraphText GetTextByName(this Document document, string name)
        {
            foreach (var text in document.GetTexts())
            {
                if (text is ParagraphText)
                {
                    if (name == text.Name)
                    {
                        return (ParagraphText)text;
                    }
                }
            }

            return null;
        }

        public static void SetCellText(this Table table, int row, int column, string text)
        {
            int rowNum = 0;
            for (uint i = 0; i < table.CellCount; i++)
            {
                var props = table.GetCellProperties(i);
                if (Math.Abs(props.TableLeftOffset - 0) < 0.0001)
                {
                    rowNum++;
                }
            }

            var cellsInRow = table.CellCount / table.RowCount;
            var index = cellsInRow * row + column;
            table.InsertText((uint)index, 0, text);

            //uint cellIndex = 
        }
    }
}
