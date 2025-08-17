using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using Filter = TFlex.DOCs.Model.Search.Filter;

namespace InventoryBook
{
    public class DataReader
    {
        private ReportParameters _reportParameters;
        private TextFormatter _tForm;

        public DataReader(ReportParameters parameters) 
        {
            this._reportParameters = parameters;
            this._tForm = _reportParameters.tForm;
        }

        public Dictionary<int, List<string>> ReadData()
        {
            var iNumberList = new List<string>();
            // Работа с объектами номенклатуры
            var refInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Nomenclature.refInfo);
            var reference = refInfo.CreateReference();
            var filter = new Filter(refInfo);
            // Объекты - только разрешенных типов
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.IsOneOf, Guids.Nomenclature.typesList);
            if (_reportParameters.iNumber)
            {
                // Список значений инвентарных номеров
                iNumberList = Enumerable.Range(_reportParameters.startNumber, _reportParameters.endNumber - _reportParameters.startNumber + 1)
                    .Select(t => t.ToString()).ToList();
                // Фильтр - все объекты, инвентарные номера которых в указанном диапазоне
                filter.Terms.AddTerm("[" + Guids.Nomenclature.BTDGroup.Fields.invnum + "]", ComparisonOperator.IsOneOf, iNumberList);
            }
            else
            {
                // Фильтр - все объекты дата выпуска которых в указанном периоде времени
                filter.Terms.AddTerm("[" + Guids.Nomenclature.BTDGroup.Fields.incomingDate + "]", ComparisonOperator.GreaterThanOrEqual, _reportParameters.startDate);
                filter.Terms.AddTerm("[" + Guids.Nomenclature.BTDGroup.Fields.incomingDate + "]", ComparisonOperator.LessThanOrEqual, _reportParameters.endDate);
            }
            // Результат поиска
            var nomenclatureList = reference.Find(filter);
            var docParams = GetNomencatureParams(nomenclatureList);

            // Работа с объектами СТО
            refInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.STO.refInfo);
            reference = refInfo.CreateReference();
            filter = new Filter(refInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, Guids.STO.type);
            if (_reportParameters.iNumber)
            {
                filter.Terms.AddTerm("[" + Guids.STO.BTDGroup.Fields.invnum + "]", ComparisonOperator.IsOneOf, iNumberList);
            }
            else
            {
                // Фильтр - все объекты дата выпуска которых в указанном периоде времени
                filter.Terms.AddTerm("[" + Guids.STO.BTDGroup.Fields.incomingDate + "]", ComparisonOperator.GreaterThanOrEqual, _reportParameters.startDate);
                filter.Terms.AddTerm("[" + Guids.STO.BTDGroup.Fields.incomingDate + "]", ComparisonOperator.LessThanOrEqual, _reportParameters.endDate);
            }
            // Результат поиска
            var stoList = reference.Find(filter);
            docParams.AddRange(GetSTOParams(stoList));

            var tableStrings = GetForm1TableStrings(docParams.OrderBy(t => t.invnum).ToList());

            return tableStrings;
        }
        // Считывание параметров документов номенклатуры
        private List<DocumentParameters> GetNomencatureParams(List<ReferenceObject> objects)
        {
            List<DocumentParameters> docParams = new List<DocumentParameters>();
            foreach (var item in objects)
            {
                var pars = new DocumentParameters();
                pars.invnum = item[Guids.Nomenclature.BTDGroup.Fields.invnum] == null ? string.Empty : item[Guids.Nomenclature.BTDGroup.Fields.invnum].ToString();
                var date = ((DateTime)item[Guids.Nomenclature.BTDGroup.Fields.incomingDate]);
                pars.incomingDate = date.Year == 1 ? String.Empty : String.Format(date.ToString("dd.MM.yy"));
                pars.denotation = item[Guids.Nomenclature.Fields.denotation] == null ? string.Empty : item[Guids.Nomenclature.Fields.denotation].ToString();
                pars.sheetsCount = item[Guids.Nomenclature.BTDGroup.Fields.sheetsCount].ToString();
                var format = item[Guids.Nomenclature.Fields.format] == null ? String.Empty : item[Guids.Nomenclature.Fields.format].ToString();
                pars.format = format;
                pars.name = item[Guids.Nomenclature.BTDGroup.Fields.name] == null ? string.Empty : item[Guids.Nomenclature.BTDGroup.Fields.name].ToString();
                var department = item[Guids.Nomenclature.BTDGroup.Fields.department] == null ? string.Empty : item[Guids.Nomenclature.BTDGroup.Fields.department].ToString();
                var depDigits = RegexPatterns.digits.Matches(department);
                // Подразделение (выделение цифровой части из названия)
                pars.department = depDigits.Count > 0 ? depDigits[0].Value : department;
                var comment = item[Guids.Nomenclature.BTDGroup.Fields.comment] == null ? string.Empty : item[Guids.Nomenclature.BTDGroup.Fields.comment].ToString();
                // Если формат - составной, то читается из примечания спецификации и пишется в примечание
                pars.comment = format.Contains("*") ? GetSpecRemark(item) + "; " + comment : comment;
                docParams.Add(pars);
            }
            return docParams;
        }
        // Считывание параметров документов СТО
        private List<DocumentParameters> GetSTOParams(List<ReferenceObject> objects)
        {
            List<DocumentParameters> docParams = new List<DocumentParameters>();
            foreach (var item in objects)
            {
                var pars = new DocumentParameters();
                pars.invnum = item[Guids.STO.BTDGroup.Fields.invnum] == null ? string.Empty : item[Guids.STO.BTDGroup.Fields.invnum].ToString();
                var date = ((DateTime)item[Guids.STO.BTDGroup.Fields.incomingDate]);
                pars.incomingDate = date.Year == 1 ? String.Empty : String.Format(date.ToString("dd.MM.yy"));
                pars.denotation = item[Guids.STO.Fields.denotation] == null ? string.Empty : item[Guids.STO.Fields.denotation].ToString();
                pars.sheetsCount = item[Guids.STO.BTDGroup.Fields.sheetsCount].ToString();
                pars.format = "A4";
                pars.name = item[Guids.STO.Fields.name] == null ? string.Empty : item[Guids.STO.Fields.name].ToString();
                var department = item[Guids.STO.BTDGroup.Fields.department] == null ? string.Empty : item[Guids.STO.BTDGroup.Fields.department].ToString();
                var depDigits = RegexPatterns.digits.Matches(department);
                // Подразделение (выделение цифровой части из названия)
                pars.department = depDigits.Count > 0 ? depDigits[0].Value : department;
                var comment = item[Guids.STO.BTDGroup.Fields.comment] == null ? string.Empty : item[Guids.STO.BTDGroup.Fields.comment].ToString();
                pars.comment = comment;
                docParams.Add(pars);
            }
            return docParams;
        }

        private Dictionary<int, List<string>> GetForm1TableStrings(List<DocumentParameters> docParams)
        {
            var tableStrings = new Dictionary<int, List<string>>();
            // Список из пустых строк
            for (int i = 0; i < GetListParamsCount(docParams); i++)
            {
                var tableString = new List<string> { String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty };
                tableStrings.Add(i, tableString);
            }

            // Добиваем из списка docParams
            int cnt = 0;
            foreach (var item in docParams)
            {
                var currentString = tableStrings[cnt];
                currentString[0] = item.invnum;
                currentString[1] = item.incomingDate;
                currentString[2] = item.denotation;
                currentString[3] = item.sheetsCount;
                currentString[4] = item.format;
                currentString[6] = item.department;
                // Обработка наименования (многострочный текст)
                var strsName = GetContentList(item.name, 60f);
                // Обработка примечания (многострочный текст)
                var strsComment = GetContentList(item.comment, 25f);
                // Заполнение параметра "Наименование"
                if (strsName.Count() > 0)
                {
                    cnt--;
                    foreach (var str in strsName)
                    {
                        cnt++;
                        currentString = tableStrings[cnt];
                        currentString[5] = str;
                    }
                    // Откатить назад на количество строк "Наименование"
                    cnt -= (strsName.Count() - 1);
                }
                // Заполнение параметра "Примечание"
                if (strsComment.Count() > 0)
                {
                    cnt--;
                    foreach (var str in strsComment)
                    {
                        cnt++;
                        currentString = tableStrings[cnt];
                        currentString[8] = str;
                    }
                    // Откатить назад на количество строк "Примечание"
                    cnt -= (strsComment.Count() - 1);
                }
                var maxCount = Math.Max(strsName.Count(), strsComment.Count());
                // Увеличить на максимальное количество строк из "Наименование" и "Примечание"
                if (maxCount > 0)
                    cnt += Math.Max(strsName.Count(), strsComment.Count()) - 1;
                // Переход на следующую строку
                cnt++;
            }

            return tableStrings;
        }

        // Значение примечания спецификации (для тех, у кого формат - составной и указан в примечании)
        private string GetSpecRemark(ReferenceObject currentObject)
        {
            string result = String.Empty;
            // Любой родительский объект
            var parent = currentObject.Parents.FirstOrDefault();

            if (parent != null)
            {
                // Иерархическая ссылка на currentObject
                var hlink = parent.GetChildLink(currentObject) as NomenclatureHierarchyLink;
                // Примечание спецификации
                result = String.IsNullOrEmpty(hlink.Remarks) ? String.Empty : hlink.Remarks.ToString().Replace("*", "").Replace(")", "");
            }

            return result;
        }

        // Разбиение длинной строки на подстроки
        private string[] GetContentList(string content, double maxWidth)
        {
            return _tForm.Wrap(content, maxWidth);
        }
        // Сколько строк в списке (с учетом разбиения на строки значения параметра "Наименование", "Примечание")
        private int GetListParamsCount(List<DocumentParameters> docParams)
        {
            int cnt = 0;
            foreach (var item in docParams)
            {
                int maxStrings = Math.Max(GetContentList(item.name, 60f).Count(), GetContentList(item.comment, 25f).Count());
                if (maxStrings > 1)
                    cnt += maxStrings;
                else
                    cnt++;
            }
            return cnt;
        }
    }
}
