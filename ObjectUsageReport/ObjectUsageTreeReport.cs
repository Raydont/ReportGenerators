using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References;
using System.IO;
using System;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model;
using ReportHelpers;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ExplorerBar;

namespace ObjectUsageReport
{
    public class ObjectUsageTreeReport
    {
        public static void Generate(MacroProvider macro)
        {
            if (macro.ТекущийОбъект == null)
            {
                macro.Сообщение("Внимание!", "Не выбран объект номенклатуры!");
                return;
            }

            var rOps = new ReportOperations(macro);
            List<ObjectHierarchyElement> tree = new List<ObjectHierarchyElement>();
            // Начальный элемент стека
            tree.Add(new ObjectHierarchyElement { level = 0, nomObject = macro.ТекущийОбъект, childObject = null });
            rOps.GetObjectsTree(macro.ТекущийОбъект, tree);
            // Получение всех ветвей входимости объекта
            var hCollection = rOps.GetHierarchyList(tree);
            var reportParams = rOps.GetReportParameters(hCollection);

            var exOps = new ExcelOperations(macro);
            exOps.MakeReport(reportParams);
        }
    }

    public class ReportOperations
    {
        MacroProvider macro;
        public ReportOperations(MacroProvider macro)
        {
            this.macro = macro;
        }
        public void GetObjectsTree(Объект currentObject, List<ObjectHierarchyElement> tree)
        {
            int lvl = 0;
            foreach (var item in currentObject.РодительскиеОбъекты)
            {
                try
                {
                    var foundObject = tree.FirstOrDefault(t => t.nomObject == currentObject);
                    if (foundObject != null) lvl = foundObject.level;
                    tree.Add(new ObjectHierarchyElement { level = lvl + 1, nomObject = item, childObject = currentObject});
                }
                catch
                {
                    continue;
                }
                GetObjectsTree(item, tree);
            }
        }

        public Dictionary<int, List<ObjectHierarchyElement>> GetHierarchyList(List<ObjectHierarchyElement> tree)
        {
            var stack = new List<ObjectHierarchyElement>();
            var chain = new List<ObjectHierarchyElement>();
            var output = new Dictionary<int, List<ObjectHierarchyElement>>();
            // Начальное условие
            stack.Add(tree[0]);
            int chainNumber = 0;
            while (stack.Count > 0)
            {
                // Верхний элемент стека переносится в список chain, затем удаляется из стека
                var currentObject = stack[stack.Count - 1];
                int level = currentObject.level;
                chain.Add(currentObject);
                stack.Remove(currentObject);
                // Находятся все родители текущего объекта (все, у которых дочерние объекты - текущий объект)
                var parents = tree.Where(t => t.childObject == currentObject.nomObject).ToList();
                if (parents.Count > 0)
                {
                    // Если родители у объекта есть, то они заносятся в стек
                    stack.AddRange(parents);
                }
                else
                {
                    // В противном случае список chain передается на выход, и из него удаляются все объекты,
                    // уровень которых выше либо равен уровню текущего объекта стека
                    chainNumber++;
                    output.Add(chainNumber, chain);
                    if (stack.Count > 0)
                    {
                        level = stack[stack.Count - 1].level;
                        chain = chain.Where(t => t.level < level).ToList();
                    }
                }
            }
            return output;
        }

        public Dictionary<int, List<NomenclatureParameters>> GetReportParameters(Dictionary<int, List<ObjectHierarchyElement>> hCollection)
        {
            // Сортировка ветвей по обозначению последнего элемента списка
            var hList = hCollection.ToList();
            hList.Sort((t1, t2) => t1.Value[t1.Value.Count - 1].nomObject[Guids.Nomenclature.Fields.denotation.ToString()].ToString().
            CompareTo(t2.Value[t2.Value.Count - 1].nomObject[Guids.Nomenclature.Fields.denotation.ToString()].ToString()));
            var newhCollection = hList.ToDictionary(t => t.Key, t => t.Value);

            var paramCollection = new Dictionary<int, List<NomenclatureParameters>>();

            foreach (var item in newhCollection)
            {
                List<NomenclatureParameters> paramList = new List<NomenclatureParameters>();
                foreach (var value in item.Value)
                {
                    var nomObject = value.nomObject;
                    var nomParams = new NomenclatureParameters();
                    nomParams.objectID = nomObject["ID"];
                    nomParams.name = nomObject[Guids.Nomenclature.Fields.name.ToString()];
                    nomParams.denotation = nomObject[Guids.Nomenclature.Fields.denotation.ToString()];
                    nomParams.code = nomObject[Guids.Nomenclature.Fields.code.ToString()];
                    nomParams.firstUsage = nomObject[Guids.Nomenclature.Fields.firstUsage.ToString()];
                    nomParams.type = nomObject.Тип.ToString();
                    nomParams.level = value.level;
                    // Количество через NomenclatureHierarchyLink
                    if (value.childObject == null)
                        nomParams.amount = 1;
                    else
                    {
                        // Иерархическая ссылка на объект childObject
                        var hlink = ((ReferenceObject)nomObject).GetChildLink((ReferenceObject)value.childObject) as NomenclatureHierarchyLink;
                        nomParams.amount = hlink.Amount;
                    }
                    paramList.Add(nomParams);                   
                }
                paramCollection.Add(item.Key, paramList);
            }
            return paramCollection;
        }
    }

    public class ExcelOperations
    {
        MacroProvider macro;
        public ExcelOperations(MacroProvider macro)
        {
            this.macro = macro;
        }

        public void MakeReport(Dictionary<int, List<NomenclatureParameters>> hCollection)
        {
            bool isCatch = false;
            // Создание копии шаблона
            var reference = new FileReference(ServerGateway.Connection);
            var pattern = reference.FindByPath(@"Служебные файлы\Шаблоны отчётов\Отчеты Excel\Excel.xlsx");
            string appDataFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Ведомость применяемости.xlsx");
            // Перезапись файла шаблона в случае его отсутствия или изменения
            if (!File.Exists(pattern.LocalPath) || File.GetLastWriteTime(pattern.LocalPath) != pattern.SystemFields.EditDate)
            {
                pattern.GetHeadRevision();
            }
            // Открытие нового документа
            Xls xls = new Xls();

            try
            {
                // формирование заголовков отчета сводной ведомости
                xls[1, 3].SetValue("ID");
                xls[1, 3].ColumnWidth = 10.0;
                xls[1, 3].Font.Bold = true;
                xls[2, 3].SetValue("Наименование");
                xls[2, 3].ColumnWidth = 70.0;
                xls[2, 3].Font.Bold = true;
                xls[3, 3].SetValue("Обозначение");
                xls[3, 3].ColumnWidth = 35.0;
                xls[3, 3].Font.Bold = true;
                xls[4, 3].SetValue("Шифр");
                xls[4, 3].ColumnWidth = 35.0;
                xls[4, 3].Font.Bold = true;
                xls[5, 3].SetValue("Тип объекта");
                xls[5, 3].ColumnWidth = 25.0;
                xls[5, 3].Font.Bold = true;
                xls[6, 3].SetValue("К-во объекта");
                xls[6, 3].ColumnWidth = 15.0;
                xls[6, 3].Font.Bold = true;

                var reportObject = hCollection[1][0];
                string mainHeader = "Применяемость объекта " + reportObject.name + "  " + reportObject.denotation;
                string firstUsage = reportObject.firstUsage;
                xls[2, 1].SetValue(mainHeader);
                xls[2, 1].Font.Bold = true;
                xls[2, 1].Font.Underline = true;
                xls[2, 1].Font.Size = 14;
                xls[2, 1, 3, 1].Merge();
                xls[2, 1].WrapText = true;

                xls[1, 3, 6, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

                int row = 3;                     // Текущая строка
                int rownach = 0;                 // Начальная строка каждой ветви
                int cnt = 0;                     // Счетчик ветвей
                macro.ДиалогОжидания.Показать("Идет построение дерева объекта " + reportObject.name + " " + reportObject.denotation, true);

                foreach (var item in hCollection)
                {
                    var list = item.Value;
                    var lastObject = list[list.Count - 1];

                    string headerCode = String.Empty;
                    row++;
                    int groupRow = row + 1;      // Строка группировки
                    int headerRow = row;         // Строка заголовка группировки
                    double previousAmount = 1;
                    cnt++;

                    foreach (var parameters in list)
                    {
                        rownach = row;
                        row++;
                        xls[1, row].SetValue(parameters.objectID);
                        // Если первичная применяемость находится среди элементов ветки, то она выделяется синим цветом
                        if (!String.IsNullOrEmpty(firstUsage) && (firstUsage == parameters.denotation))
                            xls[2, row, 4, 1].Font.Color = Convert.ToInt32(0xFF0000);
                        xls[2, row].SetValue(parameters.name);
                        xls[2, row].WrapText = true;
                        xls[3, row].SetValue(parameters.denotation);
                        xls[3, row].WrapText = true;
                        xls[4, row].SetValue(parameters.code);
                        xls[4, row].WrapText = true;
                        // Если шифр устройства отсутсвует, то в параметр Заголовок пишется его обозначение
                        if (String.IsNullOrEmpty(parameters.code))
                            headerCode = parameters.denotation;
                        else
                            headerCode = parameters.code;
                        xls[5, row].SetValue(parameters.type);
                        // Количество считается нарастающим итогом (текущее количество * количество в предыдущей итерации)
                        var newAmount = parameters.amount * previousAmount;
                        xls[6, row].SetValue(newAmount);
                        previousAmount = newAmount;
                        xls[1, rownach, 6, row - rownach + 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                      XlBordersIndex.xlEdgeBottom,
                                                                                                                      XlBordersIndex.xlEdgeLeft,
                                                                                                                      XlBordersIndex.xlEdgeRight,
                                                                                                                      XlBordersIndex.xlInsideVertical);
                        xls[2, row].Font.Bold = true;
                        xls[2, row].Font.Underline = true;
                    }

                    // Группировка данных
                    xls[1, groupRow, 6, row - groupRow + 1].Rows.Group();
                    xls[1, groupRow, 6, row - groupRow + 1].EntireRow.Hidden = true;
                    // Группировка сверху-вниз
                    xls.Worksheet.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;

                    string header = "Ветвь " + cnt + " - в составе " + headerCode;
                    xls[2, headerRow].Font.Bold = true;
                    xls[2, headerRow].SetValue(header);
                    xls[1, headerRow, 6, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    row++;

                    double percent = hCollection.Count == 0 ? 0 : (double)item.Key / hCollection.Count * 100;
                    if (!macro.ДиалогОжидания.СледующийШаг("Построение ветви в составе " + headerCode, percent))
                    {
                        macro.Сообщение("Внимание!", "Отменено пользователем");
                        isCatch = true;
                    }
                    if (hCollection.Count < 5)
                        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                }
                macro.ДиалогОжидания.Скрыть();
                xls[1, 1, 6, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                // Сохранение файла
                if (File.Exists(appDataFileName))
                    File.Delete(appDataFileName);
                xls.SaveAs(appDataFileName);
            }
            catch (Exception ex)
            {
                macro.Сообщение("Исключительная ситуация!", ex.Message);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (isCatch)
                    xls.Quit(false);
                else
                {
                    xls.Quit(true);
                    // Открытие файла
                    Process.Start(appDataFileName);
                }
            }
        }
    }

    // Класс для получения дерева иерархии 
    public class ObjectHierarchyElement
    {
        public int level;                       // Уровень вложенности
        public Объект nomObject;                // Объект
        public Объект childObject;              // Дочерний объект
    }
    // Класс для получения строки отчета
    public class NomenclatureParameters
    {
        public int objectID;                    // ID документа
        public string name;                     // Наименование
        public string denotation;               // Обозначение
        public string code;                     // Шифр
        public string firstUsage;               // Первичная применяемость
        public double amount;                   // Количество
        public string type;                     // Тип документа
        public int level;                       // Уровень вложенности
    }
    public static class Guids
    {
        public static class Nomenclature
        {
            public static class Fields
            {
                public static readonly Guid name = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
                public static readonly Guid denotation = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
                public static readonly Guid code = new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6");
                public static readonly Guid firstUsage = new Guid("06cb710f-e74a-43ce-944c-39e2b78f53d1");
            }
        }
    }
}
