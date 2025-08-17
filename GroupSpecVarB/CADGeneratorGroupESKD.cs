using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
//using System.Linq;
using System.Linq;
using System.Reflection;
using System.Text;
using TFlex;
using TFlex.DOCs.Common;
using TFlex.DOCs.Model;
using TFlex.Model;
using TFlex.Model.Model2D;
//using TFlex.Reporting;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Structure;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using TFlex.Reporting;
using Rectangle = TFlex.Drawing.Rectangle;

namespace GroupSpecVarB
{
    public class CADGeneratorGroupESKD : IReportGenerator
    {

        public void Generate(IReportGenerationContext context)
        {

            //Инициализация API CAD
            var applicationSessionSetup =
                new ApplicationSessionSetup
                    {
                        Enable3D = true,
                        ReadOnly = false,
                        PromptToSaveModifiedDocuments = false,
                        EnableMacros = true,
                        EnableDOCs = true,
                        DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version12,
                        ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
                    };

            if (!TFlex.Application.InitSession(applicationSessionSetup))
            {
                return;
            }
            TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.DoNotRefresh;
                //Инициализация API CAD
            context.CopyTemplateFile(); //Создаем копию шаблона
            TFlex.Model.Document document = null; //Документ CAD

            int baseDocId = GetDocsDocumentID(context); //определяем ID документа, для которого строится отчет

            //заполняем данными головной документ и формируем список дочерних документов спецификации    
            var baseDoc = new TFDDocument(baseDocId);
            var variantsList = new List<NomenclatureObject>();
            // поиск по идентификатору объекта
            // получение описания справочника          
            ReferenceInfo referenceInfo = ReferenceCatalog.FindReference(TFDGuids.RefNomenclatureGuid);
            // создание объекта для работы с данными
            Reference reference = referenceInfo.CreateReference();

            ReferenceObject baseReferenceObject = reference.Find(baseDocId);
            variantsList.AddRange(((NomenclatureObject) baseReferenceObject).GetVersions());
            var versionsHierarchy = new Dictionary<NomenclatureObject, List<HelperNomenclatureReportObject>>();
            //загружаем параметры всех исполнений
            var dictBom = new Dictionary<string, string>();
            foreach (var version in variantsList)
            {
               

                var versionHierarchy = version.Children.RecursiveLoadHierarchyLinks();
                var versionHierarchyObjects = versionHierarchy
                    .Where(t => t.ParentObjectId == version.SystemFields.Id)
                    .Select(t => new HelperNomenclatureReportObject
                                     {
                                         NomenclatureHLink = (NomenclatureHierarchyLink) t,
                                         NomenclatureObj = (NomenclatureObject) t.ChildObject,
                                         firstLevel = t.ParentObjectId == version.SystemFields.Id
                                     }).ToList();
                versionsHierarchy.Add(version, versionHierarchyObjects);

                foreach (var versH in versionHierarchyObjects)
                {
                   // MessageBox.Show(  versH.name + versH.denotation+"\n" +versH.kolichestvo);
                    try
                    {
                        dictBom.Add(versH.bomSection, versH.bomSection);
                    }
                    catch (Exception)
                    {
                    }
                   
                }
            }
          
            try
            {
                TFlex.Model.Document cadDocument = GetCadDocument(context);


                AddRowsDataToReportDoc(versionsHierarchy, cadDocument, dictBom.Count);

                Process.Start(context.ReportFilePath);

                //====================================================================================================================
                //!!! вход в форму!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                //====================================================================================================================

            
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK,
                                                     System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                if (document != null)
                    document.Close();
            }
        }

        private static ParagraphText GetTextTable(TFlex.Model.Document document, string name)
        {
            ParagraphText textTable = null;

            foreach (Text text in document.Texts)
            {
                if (text == null)
                    continue;

                if (text.Name == name)
                    textTable = text as ParagraphText;

                if (textTable != null)
                    break;
            }

            if (textTable == null)
                throw new ApplicationException(
                    "На чертеже не найден объект " + name);

            return textTable;
        }

        private static void AddNewPage(ref ParagraphText parText, ref Table dataTable, TFlex.Model.Document cadDocument)
        {
            uint columnsCount = (uint)dataTable.ColumnCount;
            var currentPage = cadDocument.Pages.Count;
            while (currentPage == cadDocument.Pages.Count)
            {
                AddLinesAfterGroup(ref parText, ref dataTable, (uint) dataTable.CellCount - columnsCount,
                                   ReportConfig.LinesBeforeHeader);
            }
        }

        private static bool CutPartNameOrDenotation(HelperNomenclatureReportObject row,
            HelperNomenclatureReportObject prewRow)
        {
            var regexTU = new Regex(@"(?:(\S+(\s)*ТУ)|(ТУ(\s)*\S+)|(\S+ТУ\S+))$");
            var regexGOST = new Regex(@"(?:((ГОСТ|ОСТ|ТУ)+[\s,\S]+))$");

            var regexNameItems = new Regex(@"\A[А-Яа-я]+(?=\s)");

            if (prewRow.bomSection == "Документация" && row.bomSection == "Документация")
            {
                var index = prewRow.denotation.LastIndexOf('-');
                if (index > 0)
                {
                    var mainMake = prewRow.denotation.Substring(0, index).Trim();
                    if (row.denotation.Contains(mainMake))
                    {
                        var whiteSimbol = "     ";
                        for (int i = 0; i < mainMake.Length; i++)
                        {
                            whiteSimbol += " ";
                        }
                        row._denotation = whiteSimbol + row.denotation.Substring(index).Trim();
                    }

                }
                if (row._denotation == null)
                    row._denotation = row.denotation;

            }

            if (prewRow.bomSection == "Прочие изделия")
            {
                row._name = row.name;
              
                if (regexTU.Match(prewRow.name).Value == regexTU.Match(row.name).Value
                    && regexNameItems.Match(prewRow.name).Value == regexNameItems.Match(row.name).Value
                    && regexNameItems.Match(row.name).Value != null && regexTU.Match(row.name).Value != null)
                {
                   row._name = row._name.Remove(0, regexNameItems.Match(row._name).Value.Length);
                    var index = row._name.IndexOf(regexTU.Match(row._name).Value);
                    
                    if (index > 0)
                    {
                        row._name = row._name.Remove(index, regexTU.Match(row._name).Value.Length);
                    }
                    return true;
                }
              
            }
       

            if (prewRow.bomSection == "Стандартные изделия")
            {
                row._name = row.name;

                if (regexGOST.Match(prewRow.name).Value == regexGOST.Match(row.name).Value
                    && regexNameItems.Match(prewRow.name).Value == regexNameItems.Match(row.name).Value
                    && regexNameItems.Match(row.name).Value != null && regexGOST.Match(row.name).Value != null)
                {
                    row._name = row._name.Remove(0, regexNameItems.Match(row._name).Value.Length);
                    var index = row._name.IndexOf(regexGOST.Match(row._name).Value);
                    
                    if (index > 0)
                    {
                        row._name = row._name.Remove(index, regexGOST.Match(row._name).Value.Length);
                    }
                    return true;

                }

            }
            return false;

        }


        private static void AddRowToParTextTable(ref ParagraphText parText, ref Table dataTable,  Dictionary<HelperNomenclatureReportObject, List<NomenclatureObject>> rowToAdd, 
            List<int> listMake,  TFlex.Model.Document cadDocument)
        {
            uint columnsCount = (uint) dataTable.ColumnCount;
            int i = 0;
            int m = 0;
           
            var writedRepObj = new HelperNomenclatureReportObject();

            AddNameGroup(ref parText, ref dataTable, rowToAdd.Keys.ToList()[0].bomSection);
            AddLinesAfterGroup(ref parText, ref dataTable, (uint) dataTable.CellCount - columnsCount,
                               ReportConfig.LinesBeforeHeader);
            var nameStringCount = 0;
            
            int countString = 0;
          
            foreach (var row in rowToAdd)
            {
                if (WriteRow(row))
                {
                    if (m > 0 && row.Key.bomSection != writedRepObj.bomSection)
                    {
                        var currentPage = cadDocument.Pages.Count;
                        if ((cadDocument.Pages.Count == 1 && countString + 1 >= 12 && i < 12) || (countString + 1 >= 20 && i > 12))
                        {
                            countString = 0;
                        }
                        else
                            AddLinesAfterGroup(ref parText, ref dataTable, (uint) dataTable.CellCount - columnsCount,
                                               ReportConfig.LinesBeforeHeader);
                        countString++;
                        if ((cadDocument.Pages.Count == 1 && countString >= 9 && i < 10) || (countString >= 15 && i > 12))
                        {
                            while (currentPage == cadDocument.Pages.Count)
                            {
                                AddLinesAfterGroup(ref parText, ref dataTable, (uint)dataTable.CellCount - columnsCount,
                                                   ReportConfig.LinesBeforeHeader);
                            }
                            AddNameGroup(ref parText, ref dataTable, row.Key.bomSection);
                            countString = 1;
                        }
                        else
                        {
                            AddNameGroup(ref parText, ref dataTable, row.Key.bomSection);
                            countString++;

                        }

                        AddLinesAfterGroup(ref parText, ref dataTable, (uint) dataTable.CellCount - columnsCount,
                                           ReportConfig.LinesBeforeHeader);
                        countString++;

                    }

                    if (parText == null || rowToAdd == null)
                        return;

                    try
                    {
                        row.Key._denotation = row.Key.denotation;
                        bool rewritedName = false;
                        if(m<rowToAdd.Keys.Count)
                        if (rowToAdd.Keys.Contains(writedRepObj))
                            rewritedName = CutPartNameOrDenotation(row.Key, writedRepObj);

                        if (rewritedName && writedRepObj._name == writedRepObj.name)
                        {
                            ReWriteRow(parText, dataTable, writedRepObj);
                        }

                        nameStringCount = InsertRow(parText, dataTable, row.Key, row.Value, listMake);
                        writedRepObj = new HelperNomenclatureReportObject();
                        writedRepObj = row.Key;
                        if ((cadDocument.Pages.Count == 1 && countString + nameStringCount >= 12 && i < 12) ||
                            (countString + nameStringCount >= 18 && i > 12))
                        {
                            countString = nameStringCount;
                        }
                        else

                            countString += nameStringCount;

                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Нельзя записать объект " + row.Key.name + " " + row.Key.denotation, "Ошибка!",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    i++;
                   
                }
                m++;
            }

        }

        private static bool WriteRow(KeyValuePair<HelperNomenclatureReportObject, List<NomenclatureObject>> row)
        {
            var regexNumberMake = new Regex(@"(?<=-)\d{1,3}");
            var writeRow = false;
            int numMake = 0;
            foreach (var repObj in row.Value)
            {
                try
                {
                    numMake = Convert.ToInt32(regexNumberMake.Match(repObj.Denotation).Value);
                }
                catch (Exception)
                {
                    numMake = 0;
                }
                if (row.Key.listNumebrMake.Contains(numMake))
                {
                    writeRow = true;
                    break;
                }
            }
         
            return writeRow;
        }

        private static void AddNameGroup(ref ParagraphText parText, ref Table dataTable, string nameGroup)
        {
            var font = GetSpecTableFont();
            var formatName = new TextFormatter(font);
            dataTable.InsertRows(1, (uint) (dataTable.CellCount - 1), Table.InsertProperties.After);
            uint newCellsCount = (uint) dataTable.CellCount;
            uint columnsCount = (uint) dataTable.ColumnCount;
            InsertNameGroupTableText(dataTable, newCellsCount - columnsCount + 4, 0, parText, formatName.WrapName(nameGroup, 80));
        }


        private static void AddLinesAfterGroup(ref ParagraphText parText, ref Table dataTable, uint u, int linesBeforeHeader)
        {
            dataTable.InsertRows(1, (uint)(dataTable.CellCount - 1), Table.InsertProperties.After);
            uint newCellsCount = (uint)dataTable.CellCount;
            uint columnsCount = (uint)dataTable.ColumnCount;
            InsertTableText(dataTable, newCellsCount - columnsCount + 4, 0, parText, "");

        }

        static int InsertRow(ParagraphText parText, Table table, HelperNomenclatureReportObject row, List<NomenclatureObject>  makes, List<int> listMake)
        {
            var regexNumberMake = new Regex(@"(?<=-)\d{1,3}");
            var font = GetSpecTableFont();
            var formatName = new TextFormatter(font);
            var notWritedListMake = new List<int>();
            var numberStr = 1;

            notWritedListMake.AddRange(listMake);
           
            table.InsertRows(1, (uint)(table.CellCount - 1), Table.InsertProperties.After);
            uint newCellsCount = (uint)table.CellCount;
            uint columnsCount = (uint)table.ColumnCount;

            InsertTableText(table, newCellsCount - columnsCount, 0, parText, row.format);
            InsertTableText(table, newCellsCount - columnsCount + 1, 0, parText, row.zona);
           

            if (row.bomSection == "Документация")
               InsertTableText(table, newCellsCount - columnsCount + 3, 0, parText, row._denotation);
            else
                InsertTableText(table, newCellsCount - columnsCount + 3, 0, parText, row.denotation);

            if ((row.bomSection == "Прочие изделия" || row.bomSection == "Стандартные изделия") && row._name != null)
            {
                InsertTableText(table, newCellsCount - columnsCount + 4, 0, parText, formatName.WrapName(row._name, 70));
                numberStr = formatName.WrapName(row._name, 70).Count();
            }
            else
            {
                InsertTableText(table, newCellsCount - columnsCount + 4, 0, parText, formatName.WrapName(row.name, 70));
                numberStr = formatName.WrapName(row.name, 70).Count();
            }

            if (row.position != 0)
            {
                var newStr = string.Empty;

                for (int i = 1; i < numberStr; i++)
                {
                    newStr += "\n";
                }
                InsertCountTableText(table, newCellsCount - columnsCount + 2, 0, parText, newStr+row.position.ToString().Trim());
            }

            foreach (var make in makes)
            {
                uint numMake;
                if (make.Version.ToString().Trim() == string.Empty)
                    numMake = 0;
                else
                {
                   numMake = Convert.ToUInt32(regexNumberMake.Match(make.Version.ToString()).Value);
                }
              
                uint column = 999;
                 try
                    {
                    for (uint k = 0; k < listMake.Count; k++)
                    {
                        if (listMake[(int)k] == numMake)
                        {
                            column = k;
                            break;
                        }
                     
                    }
                   
                //notWritedListMake.Remove((int)numMake);
                if (column != 999)
                   
                        if (row.kolichestvo != 0)

                            InsertCountTableText(table, newCellsCount - columnsCount + 5 + column, 0, parText,
                                            row.kolichestvo.ToString());
                        else
                        {

                            InsertCountTableText(table, newCellsCount - columnsCount + 5 + column, 0, parText, "Х");
                        }
                    }
                catch
                {
                  
                }
            }
            return numberStr;

        }


        static void ReWriteRow(ParagraphText parText, Table table, HelperNomenclatureReportObject row)
        {
            var font = GetSpecTableFont();
            var formatName = new TextFormatter(font);
            var regexTU = new Regex(@"(?:(\S+(\s)*ТУ)|(ТУ(\s)*\S+)|(\S+ТУ\S+))$");
            var regexGOST = new Regex(@"(?:((ГОСТ|ОСТ|ТУ)+[\s,\S]+))$");
            var regexNameItems = new Regex(@"\A[А-Яа-я]+(?=\s)");


            table.InsertRows(1, (uint) (table.CellCount - 1), Table.InsertProperties.After);
            uint newCellsCount = (uint) table.CellCount;
            uint columnsCount = (uint) table.ColumnCount;
            table.Clear(newCellsCount - 2*columnsCount + 4);
            table.Clear(newCellsCount - 2 * columnsCount + 2);
            table.DeleteRow(newCellsCount - columnsCount + 4);
            // table.InsertRows(1,newCellsCount - 3*columnsCount + 4,Table.InsertProperties.After);


            if (row.bomSection == "Прочие изделия")
            {
                row._name = row.name;

                var index = row._name.IndexOf(regexTU.Match(row._name).Value);

                if (index > 0)
                {
                    row._name = row._name.Remove(index, regexTU.Match(row._name).Value.Length);
                    row._name = row._name.Insert(regexNameItems.Match(row._name).Value.Length + 1,
                                                 regexTU.Match(row.name).Value + " ");
                }

                InsertTableText(table, newCellsCount - 2*columnsCount + 4, 0, parText,
                                formatName.WrapName(row._name, 70));
                InsetPosition(parText, table, row, formatName, newCellsCount, columnsCount);
            }
            else
            if (row.bomSection == "Стандартные изделия")
            {
                row._name = row.name;

                var index = row._name.IndexOf(regexGOST.Match(row._name).Value);

                if (index > 0)
                {
                    row._name = row._name.Remove(index, regexGOST.Match(row._name).Value.Length);
                    row._name = row._name.Insert(regexNameItems.Match(row._name).Value.Length + 1,
                                                 regexGOST.Match(row.name).Value + " ");
                }

                InsertTableText(table, newCellsCount - 2*columnsCount + 4, 0, parText,
                                formatName.WrapName(row._name, 70));

                InsetPosition(parText, table, row, formatName, newCellsCount, columnsCount);
            }
        }

        private static void InsetPosition(ParagraphText parText, Table table, HelperNomenclatureReportObject row,
                                          TextFormatter formatName, uint newCellsCount, uint columnsCount)
        {
            var numberStr = formatName.WrapName(row._name, 70).Count();

            if (row.position != 0)
            {
                var newStr = string.Empty;

                for (int i = 1; i < numberStr; i++)
                {
                    newStr += "\n";
                }
                InsertCountTableText(table, newCellsCount - 2*columnsCount + 2, 0, parText,
                                     newStr + row.position.ToString().Trim());
            }
        }

        /// Добавление форм в виде фгарментов. Заполнение переменных фрагментов.
        void AddPageForms(Dictionary<HelperNomenclatureReportObject, List<NomenclatureObject>> rowToAdd, TFlex.Model.Document cadDocument, int numberMake, uint lastPage)
        {
            
            var fileReference = new FileReference();
            string denotation = string.Empty;
            string name = string.Empty;
            foreach (var nomObjects in rowToAdd.Values.ToList())
            {
                foreach (var nomObj in nomObjects)
                {
                    if (nomObj.Version.ToString().Trim() == string.Empty)
                    {
                        denotation = nomObj.Denotation;
                        name = nomObj.Name;
                        break;
                    }
                }
            }

            string nameRegList = string.Empty;
            var listNamePage = new List<string>();
              if (GetCountMake(rowToAdd) - numberMake <= 10 )
              {
                   var pageRegList = new Page(cadDocument);
                  nameRegList = pageRegList.Name;
              }
            for (int i = (int)lastPage; i <= cadDocument.Pages.Count; i++)
            {
                listNamePage.Add("Страница "+ i);
            }


                foreach (Page page in cadDocument.Pages)
                {
                    if (listNamePage.Contains(page.Name))
                    {
                        var listNumberMakes = GetNumberMake(rowToAdd, numberMake);



                        FileLink link = null;
                        if (page.PageType != PageType.Normal)
                            continue;
                        FileReferenceObject fileObject;
                        if (page.Name == "Страница 1")
                        {
                            fileObject = fileReference.FindByRelativePath(
                                @"Служебные файлы\Шаблоны отчётов\Групповая спецификация по варианту Б (ГОСТ 2.113-75)\Групповая спецификация по варианту Б (ГОСТ 2.113-75) Первая страница.grb");
                            link = new FileLink(cadDocument, fileObject.LocalPath, true);
                        }
                        else
                        {

                            if (GetCountMake(rowToAdd) - numberMake <= 10 && page.Name == nameRegList)
                            {
                                 {
                                    fileObject =
                                        fileReference.FindByRelativePath(
                                            @"Служебные файлы\Шаблоны отчётов\Спецификация ГОСТ 2.113 (ЕСКД)\Лист регистрации изменений ГОСТ 2.503.grb");

                                    page.Rectangle = new Rectangle(0, 0, 210, 297);
                                    link = new FileLink(cadDocument, fileObject.LocalPath, true);
                                }
                            }
                            else
                            {


                                fileObject =
                                    fileReference.FindByRelativePath(
                                        @"Служебные файлы\Шаблоны отчётов\Групповая спецификация по варианту Б (ГОСТ 2.113-75)\Групповая спецификация по варианту Б (ГОСТ 2.113-75) Последующие страницы.grb");
                                link = new FileLink(cadDocument, fileObject.LocalPath, true);
                            }


                        }
                        fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
                      
                        Fragment formFragment = new Fragment(link);
                        formFragment.Page = page;

                        // задание значений переменных фрагмента
                        foreach (FragmentVariableValue var in formFragment.GetVariables())
                        {
                            switch (var.Name)
                            {
                                case "$naimen":
                                    var.TextValue = name;
                                    break;

                                case "$denotation":
                                    var.TextValue = denotation;
                                    break;
                                case "$oboznach":
                                    var.TextValue = denotation;
                                    break;
                               
                                case "$list":
                                    var.TextValue = page.Name.Remove(0, 9);
                                    break;
                                case "$i_00":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] == 0)
                                        {
                                            var.TextValue = "-";

                                        }
                                        else if (listNumberMakes[0] < 10)
                                        {
                                            var.TextValue = "0" + listNumberMakes[0].ToString();

                                        }
                                        else
                                        {
                                            var.TextValue = listNumberMakes[0].ToString();

                                        }
                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_01":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_02":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_03":
                                   if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_04":
                                   if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_05":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_06":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_07":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_08":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                                case "$i_09":
                                    if (listNumberMakes.Count > 0)
                                    {
                                        if (listNumberMakes[0] < 10)
                                            var.TextValue = "0" + listNumberMakes[0].ToString();
                                        else
                                            var.TextValue = listNumberMakes[0].ToString();

                                        listNumberMakes.RemoveAt(0);
                                    }
                                    break;
                           }


                        }
                     }

                
          
           
             }

              

        }
        string AddRegistPage( TFlex.Model.Document cadDocument)
         {
             var page = new Page(cadDocument);
             return page.Name;
           
         }


        /// Добавление форм в виде фгарментов. Заполнение переменных фрагментов.
        void AddOtherPageForms(Dictionary<HelperNomenclatureReportObject, List<NomenclatureObject>> rowToAdd, TFlex.Model.Document cadDocument, StringBuilder strB)
        {

            var fileReference = new FileReference();


            foreach (Page page in cadDocument.Pages)
            {
                FileLink link = null;
                if (page.PageType != PageType.Normal)
                    continue;
                FileReferenceObject fileObject;
                if (page.Name == "Страница 1")
                {
                    fileObject = fileReference.FindByRelativePath(
                        @"Служебные файлы\Шаблоны отчётов\Групповая спецификация по варианту Б (ГОСТ 2.113-75)\Групповая спецификация по варианту Б (ГОСТ 2.113-75) Первая страница.grb");

                    link = new FileLink(cadDocument, fileObject.LocalPath, true);
                    var formFragment = new Fragment(link);
                    formFragment.Page = page;

                    // задание значений переменных фрагмента

                    foreach (FragmentVariableValue var in formFragment.GetVariables())
                    {
                        switch (var.Name)
                        {

                            case "$OtherMakes":
                                if (strB.ToString().Trim() != string.Empty)
                                {
                                    var.TextValue = "Исполнения:    " + strB;
                                }
                                break;
                            case "$listov":
                                var.TextValue = cadDocument.Pages.Count.ToString();
                                break;


                        }

                    }
                }

            }

           

        }

        private static List<int> GetNumberMake(Dictionary<HelperNomenclatureReportObject, List<NomenclatureObject>> rowToAdd, int numaberMake)
        {
            var regexNumberMake = new Regex(@"(?<=-)\d{1,3}");
            var listNumberMakes = new List<int>();
            var listNumberMakesWithOutWrited = new List<int>();
            var countMake = 0;

            foreach (var row in rowToAdd.Values)
            {
                foreach (var obj in row)
                {
                    var reportObj = new int();
                    if (regexNumberMake.Match(obj.Denotation).Value != string.Empty)
                    {
                        try
                        {
                           
                            reportObj = Convert.ToInt32(regexNumberMake.Match(obj.Denotation).Value);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(obj.Denotation + " " + regexNumberMake.Match(obj.Denotation).Value);
                        }
                    }
                    else
                    {
                        reportObj = 0;
                    }
                    listNumberMakes.Add(reportObj);
                }
            }


            listNumberMakes = listNumberMakes.Distinct().ToList().OrderBy(t => t).ToList();


            if (listNumberMakes.Count - numaberMake <= 10)
                countMake = listNumberMakes.Count;
            else countMake = numaberMake + 10;

            for (int i = numaberMake; i < countMake; i++ )
            {
                listNumberMakesWithOutWrited.Add(listNumberMakes[i]);
            }

        return listNumberMakesWithOutWrited;
        }

        private static int GetCountMake(Dictionary<HelperNomenclatureReportObject, List<NomenclatureObject>> rowToAdd)
        {
            var regexNumberMake = new Regex(@"(?<=-)\d{1,3}");

            var listNumberMakes = new List<int>();

            foreach (var row in rowToAdd.Values)
            {
                foreach (var obj in row)
                {
                    var reportObj = new int();
                    if (regexNumberMake.Match(obj.Denotation).Value != string.Empty)
                    {
                        try
                        {

                            reportObj = Convert.ToInt32(regexNumberMake.Match(obj.Denotation).Value);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(obj.Denotation + " " + regexNumberMake.Match(obj.Denotation).Value);
                        }
                    }
                    else
                    {
                        reportObj = 0;
                    }
                    listNumberMakes.Add(reportObj);
                }
            }


            listNumberMakes = listNumberMakes.Distinct().ToList().OrderBy(t => t).ToList();



            return listNumberMakes.Count;
        }

        static Font GetSpecTableFont()
        {
            return new Font("T-FLEX Type A", 3.3f, GraphicsUnit.Millimeter);
        }
        static private uint InsertTableText(Table table, uint cell, uint pos, ParagraphText parText, string str)
        {
            uint row = pos;

           
                var text = new SpecFormCell(str);

                CharFormat ch = parText.CharacterFormat;
                if (text.Underlining == SpecFormCell.UnderliningFormat.Underlined)
                    ch.isUnderline = true;

                table.SetSelection(cell, cell);

                ParaFormat pf = parText.ParagraphFormat;
                if (text.Align == SpecFormCell.AlignFormat.Left)
                {
                    pf.HorJustification = ParaFormat.Just.Left;
                }
                if (text.Align == SpecFormCell.AlignFormat.Center)
                {
                    pf.HorJustification = ParaFormat.Just.Center;
                }
                else if (text.Align == SpecFormCell.AlignFormat.Right)
                {
                    pf.HorJustification = ParaFormat.Just.Right;
                }
                parText.ParagraphFormat = pf;

                table.InsertText(cell, row, text.Text, ch);
                row++;
            
            return row;
        }
        static private uint InsertNameGroupTableText(Table table, uint cell, uint pos, ParagraphText parText, string[] strArray)
        {
            uint row = pos;
            var str = strArray[0];

            for (int i = 1; i < strArray.Count(); i++)
            {
                str += "\n" + strArray[i];
            }

            var text = new SpecFormCell(str);

            CharFormat ch = parText.CharacterFormat;
            
                ch.isUnderline = true;

            table.SetSelection(cell, cell);

            ParaFormat pf = parText.ParagraphFormat;
       
                pf.HorJustification = ParaFormat.Just.Center;

                parText.ParagraphFormat = pf;

            table.InsertText(cell, row, text.Text, ch);
            row++;

            return row;
        }


        static private uint InsertCountTableText(Table table, uint cell, uint pos, ParagraphText parText, string text)
        {
            uint row = pos;

            CharFormat ch = parText.CharacterFormat;

            ch.isUnderline = false;
            table.SetSelection(cell, cell);

            ParaFormat pf = parText.ParagraphFormat;

            pf.HorJustification = ParaFormat.Just.Center;

            parText.ParagraphFormat = pf;

            table.InsertText(cell, row, text, ch);
            row++;

            return row;
        }

        static private uint InsertRightTableText(Table table, uint cell, uint pos, ParagraphText parText, string text)
        {
            uint row = pos;

            CharFormat ch = parText.CharacterFormat;

            ch.isUnderline = false;
            table.SetSelection(cell, cell);

            ParaFormat pf = parText.ParagraphFormat;

            pf.HorJustification = ParaFormat.Just.Right;

            parText.ParagraphFormat = pf;

            table.InsertText(cell, row, text, ch);
            row++;

            return row;
        }

        static private void InsertTableText(Table table, uint cell, uint pos, ParagraphText parText, string[] strArray)
        {

            var str = strArray[0];

            for (int i = 1; i< strArray.Count(); i++)
            {
                str += "\n" + strArray[i];
            }

            var text = new SpecFormCell(str);
           
            CharFormat ch = parText.CharacterFormat;
                if (text.Underlining == SpecFormCell.UnderliningFormat.Underlined)
                    ch.isUnderline = true;

                table.SetSelection(cell, cell);

                ParaFormat pf = parText.ParagraphFormat;
                if (text.Align == SpecFormCell.AlignFormat.Left)
                {
                    pf.HorJustification = ParaFormat.Just.Left;
                }
                if (text.Align == SpecFormCell.AlignFormat.Center)
                {
                    pf.HorJustification = ParaFormat.Just.Center;
                }
                else if (text.Align == SpecFormCell.AlignFormat.Right)
                {
                    pf.HorJustification = ParaFormat.Just.Right;
                }
                parText.ParagraphFormat = pf;

               
               
            
            table.InsertText(cell, pos, text.Text, ch);
           
        }

        /// <summary>
        /// Добавляет данные секции групповой спецификации
        /// </summary>
        /// <param name="parText">параграф-текст</param>
        /// <param name="dataTable">таблица параграф-текста</param>
        /// <param name="numTable">таблица номеров исполнений</param>
        /// <param name="sectionRows">данные секции для вставки</param>
        /// <param name="sectionNumber">порядковый номер секции</param>
        private  void AddNewSection(ref ParagraphText parText, ref Table dataTable, ref Table numTable, Dictionary<NomenclatureObject,
            List<HelperNomenclatureReportObject>> sectionRows, TFlex.Model.Document cadDocument)
        {
            if (sectionRows == null || sectionRows.Count == 0)
                return;

                var sortedRows = SortMakes(sectionRows);
                
                var strB = new StringBuilder();
                //var nameRegList = AddRegistPage(cadDocument);

                for (int k = 0; k < GetCountMake(sortedRows); k = k + 10)
                {
                    var listNumberMakes = new List<int>();

                    listNumberMakes = GetNumberMake(sortedRows, k);

                    foreach (var row in sortedRows.Keys)
                    {
                        row.listNumebrMake = new List<int>();
                        row.listNumebrMake.AddRange(listNumberMakes);
                    }
                    uint currentPage = cadDocument.Pages.Count;
                    AddRowToParTextTable(ref parText, ref dataTable, sortedRows, listNumberMakes, cadDocument);
                    AddPageForms(sortedRows, cadDocument, k, currentPage);
                 
                    if (k < GetCountMake(sortedRows) - 10)
                    {
                        AddNewPage(ref parText, ref dataTable, cadDocument);
                        if (k >= 10)
                        {
                            strB.Append(listNumberMakes[0] + "..." + listNumberMakes[listNumberMakes.Count - 1] +
                                        " - см. листы " + currentPage + "..." + (cadDocument.Pages.Count-1) + "\n");
                        }
                    }
                    else
                    {
                        if (k >= 10)
                        {
                            strB.Append(listNumberMakes[0] + "..." + listNumberMakes[listNumberMakes.Count - 1] +
                                        " - см. листы " + currentPage + "..." + (cadDocument.Pages.Count) + "\n");
                        }
                    }

                }

               
                AddOtherPageForms(sortedRows, cadDocument, strB);
        }

        public Dictionary<HelperNomenclatureReportObject, List<NomenclatureObject>> SortMakes
            (Dictionary<NomenclatureObject, List<HelperNomenclatureReportObject>> makes)
        {
            var regexNumberMake = new Regex(@"(?<=-)\d{1,3}");
            var regexNominalOtherItems = new Regex(@"(?<=-)((\d+)[.,]?\d*)(?=\s?(мкФ|пФ|Ом|кОм|МОм|/))");
            var sortedMakes = new Dictionary<HelperNomenclatureReportObject,  List<NomenclatureObject>>();
            var listObject = new List<HelperNomenclatureReportObject>();
            var listMakes = new List<NomenclatureObject>();

            foreach (var repObjects in makes.Values)
            {
                listObject.AddRange(repObjects);
            }

            FindDublicates(listObject);

            for (int i = 0; i < listObject.Count; i++)
            {
                if(regexNumberMake.Match(listObject[i].denotation).Value != string.Empty)
                {
                    try
                    {


                        listObject[i].numberMake = Convert.ToInt32(regexNumberMake.Match(listObject[i].denotation).Value);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show(listObject[i].denotation + " " + regexNumberMake.Match(listObject[i].denotation).Value);
                    }
                }
                else
                {
                    listObject[i].numberMake = 0;
                }
            }

            listObject = listObject.OrderBy(t => t.bomSection).ThenBy(t => t.position).ThenBy(t => t.numberMake).ThenBy(t => t.denotation).ThenBy(t => t.name).ToList();
           
            foreach (var repObj in listObject)

            {
                int i = 0;
                listMakes = new List<NomenclatureObject>();
                foreach (var make in makes)
                {
                    i++;

                    if (Contains(make.Value,repObj))
                    {
                        listMakes.Add(make.Key);
                    }

                    if ( (i == makes.Count))
                    {
                        sortedMakes.Add(repObj, listMakes);
                    }

                }
                
            }

            // Сортировка в соответствии с типами
            var sortedMakeByType = new Dictionary<HelperNomenclatureReportObject, List<NomenclatureObject>>();
           
            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Документация"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }
            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Детали"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }
            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Стандартные изделия"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }

            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Прочие изделия"))
            {
                obj.Key.nominalOtherItems = Convert.ToInt32(regexNominalOtherItems.Match(obj.Key.name).Value);
            }

            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Прочие изделия").OrderBy(t => t.Key.name).OrderBy(t => t.Key.nominalOtherItems))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }

            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Комплекты"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }
            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Комплексы"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }
            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Комплненты"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }
            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Компоненты (программы)"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }
            foreach (var obj in sortedMakes.Where(t => t.Key.bomSection == "Комплексы (программы)"))
            {
                sortedMakeByType.Add(obj.Key, obj.Value);
            }
         

            return sortedMakeByType;
        }



        private bool Contains(List<HelperNomenclatureReportObject> listRepObj, HelperNomenclatureReportObject nomObj)
        {
            foreach (var repObj in listRepObj)
            {
                if (repObj.name == nomObj.name && repObj.denotation == nomObj.denotation)
                    return true;
            }
            return false;
        }

        private void FindDublicates(List<HelperNomenclatureReportObject> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (String.Equals(data[mainID].name, data[slaveID].name) &&
                        String.Equals(data[mainID].denotation, data[slaveID].denotation) )
                    {
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
        }

        public  void AddRowsDataToReportDoc(
            Dictionary<NomenclatureObject, List<HelperNomenclatureReportObject>> versionsHierarchy,
            TFlex.Model.Document cadDocument, int countSection)
        {
            cadDocument.BeginChanges("Формирование отчета");
           //параграф текст в который заносятся данные
            ParagraphText dataTableParText = GetTextTable(cadDocument,Defines.ReportDocVars.DataTableParText);
            ParagraphText  numHeadTableParText = GetTextTable(cadDocument,Defines.ReportDocVars.NumHeaderParText);
                              //параграф текст в который заносятся номера исполнений
          
            // Начало редактирования чертежа
            dataTableParText.BeginEdit();
            numHeadTableParText.BeginEdit();

            // Заполнение таблицы
            var dataTable = dataTableParText.GetTableByIndex(0);
            var numHeadTable = numHeadTableParText.GetTableByIndex(0);
           
            //определяем кол-во секций в спецификацц
      
            AddNewSection(ref dataTableParText, ref dataTable, ref numHeadTable, versionsHierarchy, cadDocument); 

            cadDocument.EndChanges();

            cadDocument.Save();
            cadDocument.Close();
        }

        

public string DefaultReportFileExtension
        {
            get
            {
                return ".grb";
            }
        }
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {

            return false;
        }

         TFlex.Model.Document GetCadDocument(IReportGenerationContext context)
            {
                //Инициализация API CAD
                var applicationSessionSetup =
                    new ApplicationSessionSetup
                    {
                        Enable3D = true,
                        ReadOnly = false,
                        PromptToSaveModifiedDocuments = false,
                        EnableMacros = true,
                        EnableDOCs = true,
                        DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version12,
                        ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
                    };

                if (!TFlex.Application.InitSession(applicationSessionSetup))
                {
                    return null;
                }

                TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.DoNotRefresh;     //Инициализация API CAD               
                context.CopyTemplateFile();  //Создаем копию шаблона               
                return TFlex.Application.OpenDocument(context.ReportFilePath);
            }
        /**
             *  Получение идентификатора документа T-FLEX DOCs
             */
        int GetDocsDocumentID(IReportGenerationContext context)
        {
            int documentID;

            // Выборка Id выделенного документа в T-FlexDocs 2010
            if (context.ObjectsInfo.Count == 1) documentID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }

            return documentID;
        }

        internal static class TFDBomSection
        {
            public static readonly string Documentation = "Документация";
            public static readonly string Assembly = "Сборочные единицы";
            public static readonly string Details = "Детали";
            public static readonly string StdProducts = "Стандартные изделия";
            public static readonly string OtherProducts = "Прочие изделия";
            public static readonly string Materials = "Материалы";
            public static readonly string Komplekts = "Комплекты";
            public static readonly string Complex = "Комплексы";
            public static readonly string Components = "Компоненты";
            public static readonly string ProgramItems = "Программные изделия";

        }
        internal static class ApiDocs
        {
            public static SqlConnection GetConnection(bool open)
            {
                SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
                sqlConStringBuilder.DataSource = "S2";
                sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
                sqlConStringBuilder.Password = "reportUser";
                sqlConStringBuilder.UserID = "reportUser";
                try
                {
                    //подключение к БД T-FLEX DOCs
                    SqlConnection connection = new SqlConnection(sqlConStringBuilder.ToString());
                    SqlConnection.ClearPool(connection);
                    if (open)
                        connection.Open();

                    return connection;
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }

                return null;
            }
        }
   
    }


}
