using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;

namespace GroupSpecVarB
{
    internal static class ApiCad
    {
        private static Document _repDoc = null;

        /// <summary>
        /// Возвращает параграф-текст документа ReportDoc
        /// </summary>
        /// <param name="parTextName">имя параграф-текста</param>
        /// <returns>параграф текст</returns>
        private static ParagraphText GetParText(string parTextName)
        {
            ParagraphText ParText = ReportDoc.GetObjectByName(parTextName) as ParagraphText;

            if (ParText == null)
                throw new NullReferenceException("Не найден параграф текст с именем: " + parTextName);

            return ParText;
        }
        /// <summary>
        /// Добавляет в указанную строку знаки переноса строки
        /// </summary>
        /// <param name="parText">параграф текст</param>
        /// <param name="dataTable">таблица параграф текста</param>
        /// <param name="firstCellOfLastGroup">первая ячейка строки, в котрую происходит добавление</param>
        /// <param name="linesCount">сколько строк нужно добавить</param>
        private static void AddLinesAfterGroup(ref ParagraphText parText, ref Table dataTable, uint firstCellOfLastGroup, int linesCount)
        {
            uint OldPagesCount = ReportDoc.Pages.Count;

            for (int NewLineBeforeHeader = 0; NewLineBeforeHeader < linesCount; NewLineBeforeHeader++)
            {
                for (uint CellIndex = firstCellOfLastGroup; CellIndex < firstCellOfLastGroup + (uint)dataTable.ColumnCount; CellIndex++)
                {
                    uint CurrentCellTextLengt = (uint)(dataTable.GetCellTextLength(CellIndex));
                    dataTable.InsertText(CellIndex, CurrentCellTextLengt - 1, "\n");
                }

                //если при добавлении пустых строк произошел переход на новую страницу
                if (ReportDoc.Pages.Count > OldPagesCount)
                {
                    for (uint CellIndex = firstCellOfLastGroup; CellIndex < firstCellOfLastGroup + (uint)dataTable.ColumnCount; CellIndex++)
                    {
                        uint CurrentCellTextLengt = (uint)(dataTable.GetCellTextLength(CellIndex));
                        dataTable.SetCursorPosition(CellIndex, CurrentCellTextLengt - 1);
                        parText.Delete(1);
                    }

                    RefreshReportDoc();

                    break;
                }
            }
        }
        /// <summary>
        /// Обновляет число страниц в отчете
        /// </summary>
        private static void RefreshReportDoc()
        {
            ReportDoc.EndChanges();
            ReportDoc.BeginChanges("-");
        }
        /// <summary>
        /// Добавляет данный строки спецификации в документ ReportDoc
        /// </summary>
        /// <param name="parText">параграф текст с таблицей</param>
        /// <param name="dataTable">таблица данных в которую заносится строка данных</param>
        /// <param name="numTable">таблица номеров исполнений</param>
        /// <param name="rowToAdd">данные строки спецификации</param>
        /// <param name="firstRecord">является ли данная строка первой в спецификации</param>
        /// <param name="mustAddHeader">добавлять ли заголовок раздела спецификации</param>
        /// <param name="sectionNumber">номер секции</param>
        private static void AddRowToParTextTable(ref ParagraphText parText, ref Table dataTable, ref Table numTable, List<ReportRow> rowToAdd, bool firstRecord, bool mustAddHeader, int sectionNumber)
        {
            if (parText == null || rowToAdd == null || rowToAdd.Count == 0)
                return;

            uint OldPagesCount = _repDoc.Pages.Count;
            uint ColumnsCount = (uint)dataTable.ColumnCount;

            //вставляем пустые строки до названия раздела спецификации
            if (!firstRecord && mustAddHeader)
                AddLinesAfterGroup(ref parText, ref dataTable, (uint)dataTable.CellCount - ColumnsCount, ReportConfig.LinesBeforeHeader);

            if (!firstRecord)
                dataTable.InsertRows(1, (uint)dataTable.CellCount - 1, Table.InsertProperties.After);

            uint CellIndex = (uint)(dataTable.CellCount - ColumnsCount);

            //записываем строку данных спецификации в таблицу
            dataTable.InsertText(CellIndex, 0, rowToAdd[0].c1_Format);
            dataTable.InsertText(CellIndex + 1, 0, rowToAdd[0].c2_Zona);
            dataTable.InsertText(CellIndex + 2, 0, rowToAdd[0].c3_Position);
            dataTable.InsertText(CellIndex + 3, 0, rowToAdd[0].c4_Obozna4enie);
            dataTable.InsertText(CellIndex + 4, 0, rowToAdd[0].c5_Naimenovanie);
            dataTable.InsertText(CellIndex + 15, 0, rowToAdd[0].c7_Prime4anie);

            //добавляем кол-во для исполнений
            AddVariantsInfo(ref dataTable, CellIndex, rowToAdd);

            //добавляем заголовок раздела спецификации
          //  if (mustAddHeader)
           //     InsertSpecificationHeader(ref parText, ref dataTable, CellIndex, rowToAdd[0].Specification.BossSpecName);

            if (firstRecord || _repDoc.Pages.Count > OldPagesCount)
                FillNumHeadTable(ref numTable, sectionNumber, !firstRecord);
        }
        /// <summary>
        /// Заполняет строку данными о кол-ве для исполнений
        /// </summary>
        /// <param name="dataTable">таблица для вставки данных</param>
        /// <param name="firstRowCell">номер первой ячейки в строке, в которую добавляется информация</param>
        /// <param name="rowData">набор исполнений одного объекта</param>
        private static void AddVariantsInfo(ref Table dataTable, uint firstRowCell, List<ReportRow> rowData)
        {
            uint FirstVarCell = firstRowCell + 5;

            foreach (ReportRow variantOfRow in rowData)
            {
                //определяем порядковый номер ячейки для вставки кол-ва
                uint VarPosition = (uint)(variantOfRow.ParentDocVariantNumber % 10);

                while (VarPosition > 9)
                    VarPosition = VarPosition % 10;

                dataTable.InsertText(FirstVarCell + VarPosition, 0, variantOfRow.c6_Koli4estvo);
            }
        }
        /// <summary>
        /// Добавляет заголовок спецификации в указанную строку
        /// </summary>
        /// <param name="parText">параграф-текст</param>
        /// <param name="dataTable">таблица</param>
        /// <param name="dataRowFirstCell">первая ячейка строки, в которую происходит добавление</param>
        /// <param name="specificationName">обозначение спецификации</param>
        private static void InsertSpecificationHeader(ref ParagraphText parText, ref Table dataTable, uint dataRowFirstCell, string specificationName)
        {
            //вставляем пустые строки после обозначения спецификации
            dataTable.InsertText(dataRowFirstCell + 4, 0, new String('\n', 1));
            InsertFormatedTextIntoBeginOfCell(ref parText, ref dataTable, (int)(dataRowFirstCell + 4), String.Concat(specificationName, "\n"), true, false, true, ParaFormat.Just.Center);
            int EmptyLinesCount = 1;
            InsertEmptyLinesInRowCells(ref parText, ref dataTable, dataRowFirstCell, EmptyLinesCount, dataRowFirstCell + 4);
        }
        /// <summary>
        /// Добавляет название раздела спецификации в указанную ячейку
        /// </summary>
        /// <param name="parText">парагарф-текст</param>
        /// <param name="dataTable">таблица</param>
        /// <param name="dataRowFirstCell">первая ячейка строки, в которую происходит добавление</param>
        /// <param name="groupName">название раздела спецификации</param>
        private static void InsertGroupHeaderInDataRow(ref ParagraphText parText, ref Table dataTable, uint dataRowFirstCell, string groupName)
        {
            dataTable.InsertText(dataRowFirstCell + 4, 0, new String('\n', 2));
            InsertFormatedTextIntoBeginOfCell(ref parText, ref dataTable, (int)(dataRowFirstCell + 4), groupName, true, false, true, ParaFormat.Just.Center);
            InsertEmptyLinesInRowCells(ref parText, ref dataTable, dataRowFirstCell, 2, dataRowFirstCell + 4);
        }
        /// <summary>
        /// Добавляет строку определенного формата в начало ячейки
        /// </summary>
        /// <param name="parText">пагаграф-текст</param>
        /// <param name="dataTable">таблица</param>
        /// <param name="cell">ячейка, в которую происходит добавление</param>
        /// <param name="textToInsert">текст для вставки</param>
        /// <param name="underlined">подчеркнутый</param>
        /// <param name="bold">жирный</param>
        /// <param name="italic">наклонный</param>
        /// <param name="justification">центрирование текста</param>
        private static void InsertFormatedTextIntoBeginOfCell(ref ParagraphText parText, ref Table dataTable, int cell, string textToInsert,
            bool underlined, bool bold, bool italic, ParaFormat.Just justification)
        {
            CharFormat NewCharFormat = new CharFormat();
            NewCharFormat.isDefaultItalic = false;
            NewCharFormat.Color = parText.CharacterFormat.Color;
            NewCharFormat.FontName = parText.CharacterFormat.FontName;
            NewCharFormat.FontSize = parText.CharacterFormat.FontSize;
            NewCharFormat.isBold = bold;
            NewCharFormat.isItalic = italic;
            NewCharFormat.isStrikeout = parText.CharacterFormat.isStrikeout;
            NewCharFormat.isUnderline = underlined;
            NewCharFormat.Space = parText.CharacterFormat.Space;
            NewCharFormat.VertOffset = parText.CharacterFormat.VertOffset;
            dataTable.InsertText((uint)cell, 0, textToInsert, NewCharFormat);

            ParaFormat NewParFormat = new ParaFormat();
            NewParFormat.FirstLineOffset = parText.ParagraphFormat.FirstLineOffset;
            NewParFormat.FitOneLine = parText.ParagraphFormat.FitOneLine;
            NewParFormat.HorJustification = justification;
            NewParFormat.LeftIndent = parText.ParagraphFormat.LeftIndent;
            NewParFormat.LineSpace = parText.ParagraphFormat.LineSpace;
            NewParFormat.LineSpaceMode = parText.ParagraphFormat.LineSpaceMode;
            NewParFormat.NumberFormat = parText.ParagraphFormat.NumberFormat;
            NewParFormat.Numbering = parText.ParagraphFormat.Numbering;
            NewParFormat.NumProps = parText.ParagraphFormat.NumProps;
            NewParFormat.ReduceExtension = parText.ParagraphFormat.ReduceExtension;
            NewParFormat.RightIndent = parText.ParagraphFormat.RightIndent;
            NewParFormat.SpaceAfterLast = parText.ParagraphFormat.SpaceAfterLast;
            NewParFormat.SpaceBeforeFirst = parText.ParagraphFormat.SpaceBeforeFirst;
            NewParFormat.TabSize = parText.ParagraphFormat.TabSize;

            parText.SetSelection(new Position(0, 0, cell), new Position(textToInsert.Length, 0, cell));
            parText.ParagraphFormat = NewParFormat;
        }
        /// <summary>
        /// Добавляет пустые строки в указанную номером ячейки строку 
        /// </summary>
        /// <param name="parText">параграф-текст</param>
        /// <param name="dataTable">таблица</param>
        /// <param name="firstRowCell">номер первой в строке ячейки</param>
        /// <param name="count">число пустых строк</param>
        /// <param name="ignoredCells">набор ячеек, в которые добавлять пустые строки не нужно</param>
        private static void InsertEmptyLinesInRowCells(ref ParagraphText parText, ref Table dataTable, uint firstRowCell, int count, params uint[] ignoredCells)
        {
            string EmptyLines = new String('\n', count);

            for (uint CellNumber = firstRowCell; CellNumber < firstRowCell + (uint)dataTable.ColumnCount; CellNumber++)
            {
                bool IgnoreThisCell = false;

                foreach (uint IgnoredCell in ignoredCells)
                {
                    if (IgnoredCell == CellNumber)
                    {
                        IgnoreThisCell = true;
                        break;
                    }
                }

                if (IgnoreThisCell)
                    continue;

                dataTable.InsertText(CellNumber, 0, EmptyLines);
            }
        }
        /// <summary>
        /// Вставляет номера исполнений в таблицу исполнений
        /// </summary>
        /// <param name="numTable">таблица номеров исполнений</param>
        /// <param name="sectionNumber">порядковый номер секции</param>
        /// <param name="addNewRow">добавлять ли новую строку</param>
        private static void FillNumHeadTable(ref Table numTable, int sectionNumber, bool addNewRow)
        {
            if (sectionNumber < 0)
                throw new ArgumentException("В таблицу номеров исполнений подан неверный номер секции!");

            int startNumber = sectionNumber * 10;
            int endNumber = sectionNumber * 10 + 9;

            //вставляем новую строку в конец таблицы
            if (addNewRow)
                numTable.InsertRows(1, (uint)numTable.CellCount - 1, Table.InsertProperties.After);

            uint AddTextCellNumber = (uint)(numTable.CellCount - numTable.ColumnCount); //номер ячейки, в которую происходит вставка

            //поочередно вставляем числа интервала
            for (int CurrentNum = startNumber; CurrentNum <= endNumber; CurrentNum++)
            {
                numTable.InsertText(AddTextCellNumber, 0, GetVariantNumberString(CurrentNum)); //вставляем текст
                AddTextCellNumber++; //переходим на след. ячейку
            }
        }
        /// <summary>
        /// Возвращает строку с номером исполнения для вставки в отчет
        /// </summary>
        /// <param name="CurrentNum">номер исполнения</param>
        /// <returns>строка с порядковым номером исполнения для отчета</returns>
        private static string GetVariantNumberString(int CurrentNum)
        {
            string ZeroChar = "-"; //символ, заменющий 0
            string ResultText = null;

            //если число == 0, вставляем '-'
            //если разрядов в числе меньше 2 - вставляем в формате 0N
            //если разрядом >= 2 - вставляем как есть
            if (CurrentNum == 0)
                ResultText = ZeroChar;
            else
            {
                string CurrentNumString = CurrentNum.ToString();
                int DigitsCount = CurrentNumString.Length;

                if (DigitsCount == 1)
                    ResultText = String.Concat("0", CurrentNumString);
                else
                    ResultText = CurrentNumString;
            }

            return ResultText;
        }
        /// <summary>
        /// Добавляет в таблицу пустые строки до появления новой страницы
        /// </summary>
        /// <param name="parText">параграф-текст</param>
        /// <param name="dataTable">таблица параграф-текста</param>
        private static void AddNewTablePage(ref ParagraphText parText, ref Table dataTable)
        {
            uint LastCellIndex = (uint)(dataTable.CellCount - 1);
            uint OldPagesCount = ReportDoc.Pages.Count;

            //добавляем пустые строки до перехода на новую страницу
            do
            {
                dataTable.InsertRows(1, LastCellIndex, Table.InsertProperties.After);
            }
            while (ReportDoc.Pages.Count == OldPagesCount);

            //удаляем лишнюю строку
            dataTable.DeleteRow(LastCellIndex + 1);
            RefreshReportDoc();
        }
        /// <summary>
        /// Добавляет данные секции групповой спецификации
        /// </summary>
        /// <param name="parText">параграф-текст</param>
        /// <param name="dataTable">таблица параграф-текста</param>
        /// <param name="numTable">таблица номеров исполнений</param>
        /// <param name="sectionRows">данные секции для вставки</param>
        /// <param name="sectionNumber">порядковый номер секции</param>
        private static void AddNewSection(ref ParagraphText parText, ref Table dataTable, ref Table numTable, List<List<ReportRow>> sectionRows, int sectionNumber)
        {
            if (sectionRows == null || sectionRows.Count == 0 || sectionNumber < 0)
                return;

            //добавляем записи в таблицу параграф текста
            for (int RowIndex = 0; RowIndex < sectionRows.Count; RowIndex++)
            {
                if (sectionRows[RowIndex].Count == 0)
                    continue;

                bool AddNewSpecGroup = false;

              //  if (RowIndex == 0 || sectionRows[RowIndex][0].Specification.BossTypeID != sectionRows[RowIndex - 1][0].Specification.BossTypeID)
                    AddNewSpecGroup = true;

                AddRowToParTextTable(ref parText, ref dataTable, ref numTable, sectionRows[RowIndex], (sectionNumber == 0 && RowIndex == 0), AddNewSpecGroup, sectionNumber);
            }
        }
        /// <summary>
        /// Добавляет новые данные о номерах страниц исполнений
        /// </summary>
        /// <param name="firstSectionPageNumber">номер первой страницы текущей секции</param>
        /// <param name="lastSectionPageNumber">номер последней страницы текущей секции</param>
        /// <param name="firstVarNumber">номер первого исполнения</param>
        /// <param name="lastVarNumber">номер последнего исполнения</param>
        private static void AddSectionPageInformation(int firstSectionPageNumber, int lastSectionPageNumber, int firstVarNumber, int lastVarNumber)
        {
            Variable SectionPagesVar = FindDocVariable(ReportDoc, Defines.ReportDocVars.VariantSectionsPageNumbers);

            if (SectionPagesVar == null)
                return;

            string TextToAdd = SectionPagesVar.TextValue.Trim();

            if (TextToAdd.Length == 0)
            {
                if (firstVarNumber == lastVarNumber)
                    TextToAdd = "Исполнение ";
                else
                    TextToAdd = "Исполнения ";
            }
            else
                TextToAdd = String.Concat(TextToAdd, ",\n");

            TextToAdd = String.Concat(TextToAdd, GetIntervalString(firstVarNumber, lastVarNumber), " - см. ");

            if (firstSectionPageNumber == lastSectionPageNumber)
                TextToAdd = String.Concat(TextToAdd, "лист ", firstSectionPageNumber.ToString());
            else
                TextToAdd = String.Concat(TextToAdd, "листы ", GetIntervalString(firstSectionPageNumber, lastSectionPageNumber));

            SectionPagesVar.Expression = String.Concat("\"", TextToAdd, "\"");
        }
        /// <summary>
        /// Возвращает строку с диапазоном указанных чисел
        /// </summary>
        /// <param name="firstNumber">первое число диапазона</param>
        /// <param name="lastNumber">последнее число диапазона</param>
        /// <returns>строка диапазона</returns>
        private static string GetIntervalString(int firstNumber, int lastNumber)
        {
            if (firstNumber == lastNumber)
                return firstNumber.ToString();
            else if (lastNumber - firstNumber == 1)
                return String.Concat(firstNumber.ToString(), ", ", lastNumber.ToString());
            else
                return String.Concat(firstNumber.ToString(), "...", lastNumber.ToString());
        }
        /// <summary>
        /// Вычисляет номер первого исполнения для текущей секции
        /// </summary>
        /// <param name="sectionNumber">номер секции</param>
        /// <param name="maxVarNumber">максимальный номер исполнения</param>
        /// <returns>номер первого исполнения</returns>
        private static int GetStartVarNumber(int sectionNumber, int maxVarNumber)
        {
            if (sectionNumber * 10 > maxVarNumber)
                throw new ArgumentException("Число секций превышает число исполнений");

            return (int)sectionNumber * 10;
        }
        /// <summary>
        /// Вычисляет номер последнего исполнения для текущей секции
        /// </summary>
        /// <param name="sectionNumber">номер секции</param>
        /// <param name="maxVarNumber">максимальный номер исполнения</param>
        /// <returns>номер последнего исполнения</returns>
        private static int GetEndVarNumber(int sectionNumber, int maxVarNumber)
        {
            if (sectionNumber * 10 > maxVarNumber)
                throw new ArgumentException("Число секций превышает число исполнений");

            int MaxSectionValue = (int)sectionNumber * 10 + 9;

            if (maxVarNumber > MaxSectionValue)
                return MaxSectionValue;
            else
                return maxVarNumber;
        }
        /// <summary>
        /// Подключение документа отчета
        /// </summary>
        private static void LoadReportDoc()
        {

            _repDoc = TFlex.Application.ActiveDocument;

        }
        /// <summary>
        /// Сохраняет изменения, сделанные в ReportDoc-е
        /// </summary>
        private static void SaveDocumentChanges()
        {
           



        }
      
        /// <summary>
        /// Заполняет переменные сборки (документа шаблона)
        /// </summary>
        /// <param name="repPars">параметры отчета</param>
        private static void SetReportDocAssemblyVars(ReportConfig.ReportParams repPars)
        {
            Variable ReportDocVar = null;

            ReportDocVar = FindDocVariable(ReportDoc, "$Наименование");

            if (ReportDocVar != null)
                ReportDocVar.Expression = String.Concat("\"", repPars.Naimenovanie, "\"");

            ReportDocVar = FindDocVariable(ReportDoc, "$Обозначение");

            if (ReportDocVar != null)
                ReportDocVar.Expression = String.Concat("\"", repPars.Obozna4enie, "\"");

            ReportDocVar = FindDocVariable(ReportDoc, "$Разработчик");

            if (ReportDocVar != null)
                ReportDocVar.Expression = String.Concat("\"", repPars.Razrabot4ik, "\"");
        }


        /// <summary>
        /// Документ отчета, в котором формируется спецификация
        /// </summary>
        public static Document ReportDoc
        {
            get
            {
                if (_repDoc == null)
                    throw new NullReferenceException("Документ отчета не найден!");

                return _repDoc;
            }
        }
        /// <summary>
        /// Возвращает ID документа DOCs, для которого строится спецификация
        /// </summary>
        /// <returns>ID документа, для которого строится спецификация</returns>
        public static int GetBaseDocID()
        {

            //пытаемся получить ID обрабатываемого документа
            object attrib = ApiCad.ReportDoc.Attributes.GetTextAttribute("DOCSOBJID");

            if (attrib == null)
                return 0;

            return Int32.Parse(attrib.ToString());

        }
        /// <summary>
        /// Инициализация DOCs APi
        /// </summary>
        public static void InitApi()
        {
#if DEBUG
            TFlex.ApplicationSessionSetup setup = new ApplicationSessionSetup();
            setup.DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version10;
            setup.Enable3D = false;
            setup.EnableMacros = true;
            setup.ProtectionLicense = ApplicationSessionSetup.License.Auto;
            setup.ReadOnly = false;
            bool result = TFlex.Application.InitSession(setup);

            if (!result)
                throw new InvalidOperationException("Ошибка подключения к Cad Api");
#endif

            LoadReportDoc();
        }
        /// <summary>
        /// Возвращает переменную документа по ее названию
        /// </summary>
        /// <param name="doc">документ</param>
        /// <param name="varName">имя переменной</param>
        /// <returns>найденная переменная документа</returns>
        public static Variable FindDocVariable(Document doc, string varName)
        {
            if (doc == null || String.IsNullOrEmpty(varName))
                return null;

            foreach (Variable currentVar in doc.Variables)
            {
                if (String.Compare(currentVar.Name, varName) == 0)
                    return currentVar;
            }

            return null;
        }
        /// <summary>
        /// Формирует отчет на базе подготовленных данных
        /// </summary>
        /// <param name="repPars">параметры отчета</param>
        /// <param name="groupSpecData">данные групповой спецификации</param>
        public static void AddRowsDataToReportDoc(ReportConfig.ReportParams repPars, GroupSpecification groupSpecData)
        {
            ReportDoc.BeginChanges("Формирование спецификации");
            ParagraphText DataTableParText = GetParText(Defines.ReportDocVars.DataTableParText); //параграф текст в который заносятся данные
            ParagraphText NumHeadTableParText = GetParText(Defines.ReportDocVars.NumHeaderParText); //параграф текст в который заносятся номера исполнений
            DataTableParText.BeginEdit();
            NumHeadTableParText.BeginEdit();
            Table DataTable = DataTableParText.GetTableByIndex(0); //таблица параграф текста
            Table NumHeadTable = NumHeadTableParText.GetTableByIndex(0); //таблица заголовка номеров исполнений

#if OLD_API
            Page FirstPage = GetFirstReportPage();
            FileLink RestPagesFragmentLink = DisableNewPagesAutoFillAndGetRestPagesFileLnk(ref FirstPage);
#endif

            //определяем кол-во секций в спецификацц
            int SectionsCount = groupSpecData.GetSectionsCount();
            int OldPagesCount = 0; //номер страницы до добавления секции данных
            int NewPagesCount = 0; //номер страницы после добавления секции данных

            for (int currentSection = 0; currentSection < SectionsCount; currentSection++)
            {
                OldPagesCount = (int)_repDoc.Pages.Count + 1;

                //добавляем переход на новую страницу
                if (currentSection > 0)
                    AddNewTablePage(ref DataTableParText, ref DataTable);

                //получаем данные для текущей секции
                List<List<ReportRow>> SectionRows = groupSpecData.GetSectionRowList(currentSection);
                //добавляем данные текущей секции в форматку
                AddNewSection(ref DataTableParText, ref DataTable, ref NumHeadTable, SectionRows, currentSection);
                NewPagesCount = (int)_repDoc.Pages.Count;

                //добавляем сведения о листах, на которых расположены исполнения (после 10-го)
                if (currentSection > 0)
                {
                    int MaxVariantNumber = groupSpecData.GetMaxVariantNumber();
                    AddSectionPageInformation(OldPagesCount, NewPagesCount, GetStartVarNumber(currentSection, MaxVariantNumber), GetEndVarNumber(currentSection, MaxVariantNumber));
                }
            }

#if OLD_API
            FirstPage.FragmentForNewPageFileLink = RestPagesFragmentLink;
            FillNewPagesWithFragments(FirstPage, RestPagesFragmentLink, repPars);
#endif

            //устанавливаем переменные сборки
            SetReportDocAssemblyVars(repPars);
            NumHeadTableParText.EndEdit();
            DataTableParText.EndEdit();
            ReportDoc.EndChanges();
            SaveDocumentChanges();
        }
    }
}
