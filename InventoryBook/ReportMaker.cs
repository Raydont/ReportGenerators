using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model;
using TFlex.Model.Model2D;
using TFlex.Drawing;
using TFlex.Model;

namespace InventoryBook
{
    public class ReportMaker
    {
        TFlex.Model.Document _cadDocument;
        public ReportMaker(TFlex.Model.Document cadDocument)
        {
            this._cadDocument = cadDocument;
        }

        private ParagraphText GetDocumentTextTable(string textTableName)
        {
            ParagraphText textTable = null;

            foreach (var text in _cadDocument.GetTexts())
            {
                if (text == null)
                    continue;

                if (text.Name == textTableName)
                    textTable = text as ParagraphText;

                if (textTable != null)
                    break;
            }

            if (textTable == null)
                throw new ApplicationException(
                    "На чертеже не найден объект " + textTableName);

            return textTable;
        }

        public void FillDocument(Dictionary<int, List<string>> tableStrings)
        {
            _cadDocument.BeginChanges("Заполнение таблиц и переменных");
            // Получение таблицы документа (Форма 1)
            ParagraphText textTable = GetDocumentTextTable("Form1Table");
            textTable.BeginEdit();
            Table table = textTable.GetTableByIndex(0);
            // Добивание пустыми строками
            table.InsertRows((uint)tableStrings.Count - 1, 0, Table.InsertProperties.After);
            uint count = 0;
            // Заполнение таблицы
            foreach (var item in tableStrings)
            {
                table.InsertText(9 * count, 0, item.Value[0]);
                table.InsertText(1 + 9 * count, 0, item.Value[1]);
                table.InsertText(2 + 9 * count, 0, item.Value[2]);
                table.InsertText(3 + 9 * count, 0, item.Value[3]);
                table.InsertText(4 + 9 * count, 0, item.Value[4]);
                table.InsertText(5 + 9 * count, 0, item.Value[5]);
                table.InsertText(6 + 9 * count, 0, item.Value[6]);
                table.InsertText(8 + 9 * count, 0, item.Value[8]);
                count++;
            }
            textTable.EndEdit();

            // Наложение форматок и заполнение переменных
            var fileReference = new FileReference(ServerGateway.Connection);
            var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.Form1Gost2_501_2013);
            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
            var pageForm1Link = new FileLink(_cadDocument, fileObject.LocalPath, true);

            foreach (var page in _cadDocument.GetPages())
            {
                page.Rectangle = new Rectangle(0, 0, 210, 297);
                Fragment formFragment = new Fragment(pageForm1Link);
                formFragment.Page = page;
            }
            _cadDocument.EndChanges();
        }
    }
}
