using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.Reporting;
using TFlex.DOCs.References;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using System.Windows.Forms;
using System.Drawing;
using ReportHelpers;
using Microsoft.Office.Interop.Excel;
using TFlex.DOCs.Model.References.Macros;
using System.CodeDom;

namespace ActualComplexesReport
{
    public static class ExcelOperations
    {
        // Заполнение таблицы Excel данными, полученными из справочника
        public static int FillTable(List<ComplexParameters> paramList, bool isFull, Xls xls, MacroProvider macro)
        {
            int row = 3;
            int rowmax = 3;
            int row1;
            string cell = string.Empty;
            int cnt = 0;
            var endColumn = 19;
            if (!isFull) endColumn = 9;
            macro.ДиалогОжидания.Показать("Идет запись данных в отчет...", true);
            foreach (var parameters in paramList)
            {
                cnt++;
                xls[1, row].SetValue(cnt.ToString());
                xls[2, row].SetValue(parameters.complexName);
                xls[3, row].SetValue(parameters.denotation);
                xls[4, row].NumberFormat = "@";
                xls[4, row].SetValue(parameters.orderNumber);
                xls[5, row].SetValue(parameters.letter);
                xls[7, row].SetValue(parameters.customer);
                xls[8, row].SetValue(parameters.contractNumber);
                xls[9, row].SetValue(parameters.contractor);
                // Заполнение списка проверяемых объектов
                row1 = row;
                foreach (string checkedObject in parameters.checkedObjects)
                {
                    xls[6, row1].SetValue(checkedObject);
                    row1++;
                }
                if (row1 > rowmax)
                    rowmax = row1;

                if (isFull)
                {
                    xls[10, row].NumberFormat = "@";
                    xls[10, row].SetValue(parameters.factoryNumber);
                    xls[11, row].SetValue(parameters.dislocationPlace);
                    xls[12, row].SetValue(parameters.deliveryDate);
                    xls[13, row].SetValue(parameters.deploymentDate);
                    xls[14, row].SetValue(parameters.warranty);
                    xls[15, row].SetValue(parameters.resourceExtending);
                    // Заполнение списка контактов
                    row1 = row;
                    foreach (string contact in parameters.contractorContacts)
                    {
                        xls[16, row1].SetValue(contact);
                        row1++;
                    }
                    if (row1 > rowmax)
                        rowmax = row1;
                    // Заполнение списка бюллетеней
                    row1 = row;
                    foreach (string bulletin in parameters.bulletins)
                    {
                        xls[17, row1].SetValue(bulletin);
                        row1++;
                    }
                    if (row1 > rowmax)
                        rowmax = row1;
                    // Заполнение списка актов
                    row1 = row;
                    foreach (string act in parameters.technicalActs)
                    {
                        xls[18, row1].SetValue(act);
                        row1++;
                    }
                    if (row1 > rowmax)
                        rowmax = row1;
                    // Создание и заполнение списка примечаний
                    row1 = row;
                    if (!String.IsNullOrEmpty(parameters.comments))
                    {
                        foreach (string comment in parameters.comments.Split('\n'))
                        {
                            xls[19, row1].SetValue(comment);
                            row1++;
                        }
                    }
                    if (row1 > rowmax)
                        rowmax = row1;
                }
                // Последняя строчка для текущего комплекса
                if (rowmax > row)
                    row = rowmax;
                else
                    row++;

                for (int i = 1; i <= endColumn; i++)
                    xls[i, row].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeTop);
                double percent = paramList.Count == 0 ? 0 : (double)cnt / paramList.Count * 100;
                if (!macro.ДиалогОжидания.СледующийШаг("Строка " + row + " - записываются данные заказа №" + parameters.orderNumber, percent))
                {
                    return 0;
                }
            }
            xls[1, 3, endColumn, row - 2].CenterText();
            macro.ДиалогОжидания.Скрыть();
            return row;
        }
        // Создание вертикальных границ
        public static void CreateBorders(int row, bool isFull, Xls xls)
        {
            var endColumn = 19;
            if (!isFull) endColumn = 9;
            xls[1, 3, endColumn, row - 3].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlInsideVertical);
            xls[1, 3].Select();
            xls.Application.ActiveWindow.FreezePanes = true;

            for (int i = 3; i < row; i++)
            {
                xls[1, i, endColumn, row - i].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeLeft);
                xls[1, i, endColumn, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeRight);
            }
        }
        // Создание заголовка отчета
        public static void CreateHeader(bool isFull, Xls xls)
        {
            int col = 1;
            int row = 1;
            int col1 = 1;
            int row1 = 2;

            // Надписи в полях заголовка
            xls[col, row].Cells.ColumnWidth = 10;
            col = InsertHeader("№ п/п", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("1", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 30;
            col = InsertHeader("Наименование комплекса", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("2", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 30;
            col = InsertHeader("Обозначение", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("3", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 20;
            col = InsertHeader("Номер заказа", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("4", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 10;
            col = InsertHeader("Литера", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("5", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 30;
            col = InsertHeader("Проверяемые изделия", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("6", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 30;
            col = InsertHeader("Заказчик", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("7", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 30;
            col = InsertHeader("Номер договора", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("8", row1, col1, "Calibri", 11, xls);

            xls[col, row].Cells.ColumnWidth = 30;
            col = InsertHeader("Контрагент", row, col, "Calibri", 11, xls);
            col1 = InsertHeader("9", row1, col1, "Calibri", 11, xls);

            if (isFull)
            {
                xls[col, row].Cells.ColumnWidth = 20;
                col = InsertHeader("Заводской номер", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("10", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 30;
                col = InsertHeader("Адрес места дислокации", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("11", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 30;
                col = InsertHeader("Дата поставки", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("12", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 30;
                col = InsertHeader("Дата развертывания", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("13", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 30;
                col = InsertHeader("Срок гарантии", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("14", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 50;
                col = InsertHeader("Сведения о продлении ресурса эксплуатации", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("15", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 50;
                col = InsertHeader("Представитель эксплуатирующей организации", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("16", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 30;
                col = InsertHeader("Бюллетени", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("17", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 30;
                col = InsertHeader("Технические акты", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("18", row1, col1, "Calibri", 11, xls);

                xls[col, row].Cells.ColumnWidth = 30;
                col = InsertHeader("Примечания", row, col, "Calibri", 11, xls);
                col1 = InsertHeader("19", row1, col1, "Calibri", 11, xls);
            }
            var endColumn = 19;
            if (!isFull) endColumn = 9;
            for (int i = 1; i <= endColumn; i++)
            {
                xls[i, 1, 1, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom);
                xls[i, 1].Interior.Color = ColorToInt(Color.LightSteelBlue);
                xls[i, 2].Interior.Color = ColorToInt(Color.LightSteelBlue);
            }
        }
        // Запись в ячейку элемента заголовка таблицы
        public static int InsertHeader(string header, int row, int col, dynamic fontName, dynamic fontSize, Xls xls)
        {
            xls[col, row].Font.Name = fontName;
            xls[col, row].Font.Size = fontSize;
            xls[col, row].Font.Bold = true;
            xls[col, row].CenterText();
            xls[col, row].SetValue(header);
            xls[col, row].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                      XlBordersIndex.xlEdgeBottom,
                                                                                      XlBordersIndex.xlEdgeLeft,
                                                                                      XlBordersIndex.xlEdgeRight,
                                                                                      XlBordersIndex.xlInsideVertical);
            col++;

            return col;
        }
        // Преобразование цвета ARGB в целочисленное значение
        public static int ColorToInt(Color color)
        {
            int colorInt = 0;

            colorInt = 255 << 24 | color.B << 16 | color.G << 8 | color.R;
            return colorInt;
        }

        // Считывание данных из файла Эксель одной строкой (для импортера)
        public static string ReadMultilineString(Xls xls, int columnIndex, Dictionary<int, int> valueColumns, ComplexParameters complexParams, bool keepLines = false)
        {
            var values = new List<string>();
            // Считываются строки, посвященные конкретному комплексу, столбец связан с конкретным параметром
            for (int i = complexParams.row; i < complexParams.row + complexParams.countRows; i++)
            {
                var currentCell = xls[valueColumns[columnIndex], i];
                values.Add((currentCell.Text + "").Trim());
            }

            for (int i = 0; i < values.Count; i++)
            {
                // Зачем три пробела заменяются на один???
                while (values[i].Contains("  "))
                {
                    values[i] = values[i].Replace("  ", " ");
                }
            }
            // Удаляются все пустые строки, а также строки, содержащие "-" (т.е. прочерк)
            values.RemoveAll(string.IsNullOrEmpty);
            values.RemoveAll(t => t.Trim().All(t1 => t1 == '-'));
            // Если нужно сохранить строки, то через сброс, иначе - в одну строку (по умолчанию - в одну строку)
            if (keepLines)
            {
                return string.Join("\r\n", values).Trim();
            }
            return string.Join(" ", values).Trim();
        }
        // Считывание списка значений из файла Эксель (для импортера)
        public static List<string> ReadList(Xls xls, int columnIndex, Dictionary<int, int> valueColumns, ComplexParameters complexParams)
        {
            var values = new List<string>();
            for (int i = complexParams.row; i < complexParams.row + complexParams.countRows; i++)
            {
                var currentCell = xls[valueColumns[columnIndex], i];
                values.Add(currentCell.Text + "");
            }
            values.RemoveAll(string.IsNullOrEmpty);
            values.RemoveAll(t => t.Trim().All(t1 => t1 == '-'));
            return values;
        }
    }

    // Индексы строки комплекса (из файла Эксель для импорта)
    public static class FileString
    {
        public static int ItemIndex = 1;                         // Номер пункта
        public static int NameIndex = 2;                         // Наименование комплекса
        public static int DenotationIndex = 3;                   // Обозначение
        public static int OrderNumberIndex = 4;                  // Номер заказа
        public static int LetterIndex = 5;                       // Литера
        public static int FactoryNumberIndex = 6;                // Заводской номер
        public static int CheckedObjectsIndex = 7;               // Проверяемые изделия (список)
        public static int CustomerIndex = 8;                     // Заказчик
        public static int DislocationPlaceIndex = 9;             // Адрес поставки (мультистрочный)
        public static int DeliveryDateIndex = 10;                // Дата поставки
        public static int DeploymentDateIndex = 11;              // Дата развертывания        
        public static int ContractNumberIndex = 12;              // Номер договора
        public static int ContractorIndex = 13;                  // Контрагент (мультистрочный)
        public static int WarrantyIndex = 14;                    // Срок гарантии
        public static int ContractorContactIndex = 15;           // Представитель эксплуатирующей организации (список)
        public static int BulletinsIndex = 16;                   // Бюллетени (список)
        public static int TechnicalActsIndex = 17;               // Технические акты (список)
        public static int CommentsIndex = 18;                    // Примечания (мультистрочный)
        public static int ResourceExtendingIndex = 19;            // Сведения о продлении ресурса эксплуатации (мультистрочый)        	
    }
    public class ComplexParameters
    {
        public int row;                                                 // номер строки (для импортера)
        public int countRows;                                           // количество строк (для импортера)
        public string number;                                           // номер позиции
        public string complexName;                                      // наименование комплекса
        public string denotation;                                       // обозначение комплекса
        public string orderNumber;                                      // номер заказа
        public string letter;                                           // литера
        public string factoryNumber;                                    // заводской номер
        public List<string> checkedObjects = new List<string>();        // проверяемые изделия
        public string customer;                                         // заказчик
        public string contractNumber;                                   // номер заказа
        public string contractor;                                       // контрагент
        public string dislocationPlace;                                 // место дислокации
        public string deliveryDate;                                     // дата поставки
        public string deploymentDate;                                   // дата развертывания
        public string warranty;                                         // срок гарантии
        public string resourceExtending;                                // продление ресурса эксплуатации
        public List<string> contractorContacts = new List<string>();    // представитель эксплуатирующей организации
        public List<string> bulletins = new List<string>();             // бюллетени
        public List<string> technicalActs = new List<string>();         // технические акты
        public string comments;                                         // примечания
    }

    public static class FormParameters
    {
        // Порядок записи условий для фильтра
        public enum StringTermsIndexes
        {
            Exploitation = 0,
            Customer = 1,
            CheckedObjects = 2,
            Structures = 3,
            Actual = 4,
            Cancelled = 5
        }
        // Названия полей формы отчета
        public static class FormNames
        {
            public static string Device = "Аппаратура";
            public static string Denotation = "Обозначение";
            public static string Order = "Заказ";
            public static string Customer = "Потребитель";
            public static string Exploitation = "Эксплуатация";
            public static string CheckedObjects = "Список изделий";
            public static string Structures = "Структуры";
            public static string Actual = "Актуальность";
            public static string IsFull = "Расширенный вид";
            public static string Cancelled = "Списание";
        }
        // Значения полей формы отчета
        public static class FormValues
        {
            public static class Exploitation
            {
                public static string True = "Эксплуатируется";
                public static string False = "Не эксплуатируется";
                public static string All = "Все заказы";
            }
            public static class Structures
            {
                public static string True = "Есть файлы структур";
                public static string False = "Нет файлов структур";
                public static string All = "Все заказы";
            }
            public static class Cancelled
            {
                public static string True = "Списано";
                public static string False = "Не списано";
                public static string All = "Все комплексы";
            }
            public static class Customer
            {
                public static string All = "Все типы";
            }
            public static class Actual
            {
                public static string True = "С рабочими экземплярами";
                public static string False = "Без рабочих экземпляров";
                public static string All = "Все заказы";
            }
            public static class CheckedObjects
            {
                public static string True = "Есть проверяемые изделия";
                public static string False = "Нет проверяемых изделий";
                public static string All = "Все заказы";
            }
        }       
    }
    public static class Guids
    {
        public static class OrderStructureReference
        {
            public static readonly Guid refGuid = new Guid("96c2b26e-4217-47d8-b344-969054feaaae");
            public static readonly Guid type = new Guid("4c6c582b-d4d0-4815-b0ff-bdbec80c7e8e");
            public static class Fields
            {
                public static readonly Guid orderNumber = new Guid("027fcfd8-4bb6-4b09-9423-7f2ffc0ae0dc");
                public static readonly Guid exploitation = new Guid("bd9e0c3b-df7a-464c-a422-41084196a104");
                public static readonly Guid customer = new Guid("c94c3101-4c74-4c30-8454-be056482cdff");
                public static readonly Guid contractNumber = new Guid("67722dd8-97e5-4669-9510-c7539f8f3be6");
                public static readonly Guid contractor = new Guid("3ed0fe36-7dab-4e8c-b005-a9d45aaa5e9b");
            }
            public static class Links
            {
                public static readonly Guid complexes = new Guid("f22aa93e-a7cc-4938-9d9b-3d9f90718697");
                public static readonly Guid structures = new Guid("eb6b67f0-03ba-406a-88dd-bdf1d45cc864");
                public static readonly Guid checkedObjects = new Guid("9208b2e5-523e-43b3-b2cb-4b372981465c");
                public static readonly Guid actualComplexes = new Guid("49ee09b0-8259-4ba5-86e4-5a2cf2aa1171");
            }
            public static class CheckedObjects
            {
                public static readonly Guid type = new Guid("8f077aef-9da5-4bf3-a73d-b2f8c524941b");
                public static readonly Guid name = new Guid("865a0fdb-33ab-41a1-b28c-a1cae40aa134");
            }
        }
        public static class ComplexReference
        {
            public static readonly Guid refGuid = new Guid("e876b55d-27e2-4a79-841a-e77a86fd5c85");
            public static readonly Guid type = new Guid("719f3d83-0448-4f7c-9ec5-ce66d69b128a");
            public static class Fields
            {
                public static readonly Guid name = new Guid("7a70c6f4-f586-4739-b84a-cadef6454f49");
                public static readonly Guid denotation = new Guid("bc5f447c-8e7e-4cd1-b43c-da5920e9019a");
                public static readonly Guid letter = new Guid("2878ac40-41cb-4410-b17f-6209c9214c9c");
            }
        }
        public static class ActualComplexReference
        {
            public static readonly Guid refGuid = new Guid("f08f91fe-8b29-4a27-9ca4-df8e8001fb98");
            public static readonly Guid type = new Guid("a1522f9d-6fab-43da-8c75-c357bb48b35a");
            public static class Fields
            {
                public static readonly Guid number = new Guid("a08ccd18-62e9-4241-9e26-356c42573eab");
                public static readonly Guid dislocationPlace = new Guid("e1ad2470-0fa0-4791-8ef5-345cfea13d62");
                public static readonly Guid deliveryDate = new Guid("353aeba5-89ca-42d1-906c-d38acc0ce8e2");
                public static readonly Guid deploymentDate = new Guid("0cd51111-7352-4f46-bd03-3872f8669a7a");
                public static readonly Guid warranty = new Guid("ef778c72-884f-47f7-9957-f8e4dcae107e");
                public static readonly Guid resourceExtending = new Guid("bdd97747-1159-412b-a0cb-c9f87c7c2b3a");
                public static readonly Guid cancelled = new Guid("80e54115-dc5a-42de-9ba1-66bab10c2fe9");
                public static readonly Guid comments = new Guid("d5546242-db50-4715-8a8e-19c140064783");
            }
            public static class Links
            {
                public static readonly Guid order = new Guid("49ee09b0-8259-4ba5-86e4-5a2cf2aa1171");
                public static readonly Guid bulletins = new Guid("dfe683da-279c-40b1-bbc9-ffa62b675331");
                public static readonly Guid contractorContacts = new Guid("8abcc585-5637-428b-8568-f04c905249ca");
                public static readonly Guid technicalActs = new Guid("737fb282-0c88-48b4-bc27-43481985aa88");
            }
            public static class Bulletins
            {
                public static readonly Guid type = new Guid("2a51e505-6e42-4f06-b70d-133f44b6288a");
                public static readonly Guid name = new Guid("f7668409-740f-4028-a6b6-415e4292e09f");
            }
            public static class ContractorContacts
            {
                public static readonly Guid type = new Guid("e486bc98-719a-4fd3-adaf-9a2abd09664f");
                public static readonly Guid name = new Guid("58780263-32f5-4a96-b2e4-15f43bf14872");
            }
            public static class TechnicalActs
            {
                public static readonly Guid type = new Guid("31658a1b-8de7-4087-aa79-bdf542099b0c");
                public static readonly Guid name = new Guid("ff2ab928-1049-47d0-b87b-d9c8f1a95593");
            }
        }
    }
}
