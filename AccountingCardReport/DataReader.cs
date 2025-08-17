using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Drawing;
using TFlex.Model.Measure;
using static TFlex.Model.Model2D.FileBrowserDialog;

namespace AccountingCardReport
{
    public class DataReader
    {
        NomenclatureReference _nomReference;
        TextFormatter _tForm;
        // Конструктор для чтения данных из объекта справочника (кроме "Номенклатура")
        public DataReader(TextFormatter tForm)
        {
            this._tForm = tForm;
        }
        // Конструктор для чтения из справочника "Номенклатура" (сам справочник нужен для поиска объекта первичной примеяемости)
        public DataReader(NomenclatureReference nomReference, TextFormatter tForm) 
        { 
            this._nomReference = nomReference;
            this._tForm = tForm;
        }
        // Чтение данных документа из справочника "Номенклатура"
        public DocumentParameters GetDocumentData(ReferenceObject currentObject)
        {
            List<NomenclatureObject> makes = new List<NomenclatureObject>();
            // Исполнения документа (начиная с базового исполнения)
            /*makes = ((NomenclatureObject)currentObject).GetVersions().
                OrderByDescending(t => (t.Version.IsNull || t.Version.IsEmpty) ? 0 : (int)t.Version).ToList();*/
            makes = ((NomenclatureObject)currentObject).GetVersions().OrderBy(t => GetMakeNumber(t)).ToList();

            if (makes.Count == 0)
            {
                // Если у объекта исполнений нет, то в список добавляется только он один
                makes.Add((NomenclatureObject)currentObject);
            }
            else
            {
                // Текущий объект - всегда основное исполнение
                currentObject = (ReferenceObject)makes.First();
            }
            var lastMakeNumberString = GetMakeNumberString(makes.Last());
            // Максимальный номер исполнения 
            string maxMake = (lastMakeNumberString == string.Empty) ? string.Empty : " (-" + lastMakeNumberString + ")";
            // Заполнение параметров текущего объекта
            var pars = new DocumentParameters();
            pars.name = currentObject[Guids.Nomenclature.BTDGroup.Fields.name] == null ? string.Empty : currentObject[Guids.Nomenclature.BTDGroup.Fields.name].ToString();
            var denotation = currentObject[Guids.Nomenclature.Fields.denotation] == null ? string.Empty : currentObject[Guids.Nomenclature.Fields.denotation].ToString();
            // К обозначению добавляется максимальный номер исполнения
            pars.denotation = denotation + maxMake;
            pars.doctype = currentObject[Guids.Nomenclature.BTDGroup.Fields.doctype] == null ? string.Empty : currentObject[Guids.Nomenclature.BTDGroup.Fields.doctype].ToString();
            var department = currentObject[Guids.Nomenclature.BTDGroup.Fields.department] == null ? string.Empty : currentObject[Guids.Nomenclature.BTDGroup.Fields.department].ToString();
            var depDigits = RegexPatterns.digits.Matches(department);
            // Подразделение (выделение цифровой части из названия)
            pars.department = depDigits.Count > 0 ? depDigits[0].Value : department;
            pars.invnum = currentObject[Guids.Nomenclature.BTDGroup.Fields.invnum] == null ? string.Empty : currentObject[Guids.Nomenclature.BTDGroup.Fields.invnum].ToString();
            pars.manufactoryCode = currentObject[Guids.Nomenclature.BTDGroup.Fields.manufactoryCode] == null ? string.Empty : currentObject[Guids.Nomenclature.BTDGroup.Fields.manufactoryCode].ToString();
            pars.original = pars.manufactoryCode;
            pars.sheetsCount = currentObject[Guids.Nomenclature.BTDGroup.Fields.sheetsCount].ToString();
            var date = ((DateTime)currentObject[Guids.Nomenclature.BTDGroup.Fields.incomingDate]);
            pars.incomingDate = date.Year == 1 ? String.Empty : String.Format(date.ToString("dd.MM.yy"));
            var format = currentObject[Guids.Nomenclature.Fields.format] == null ? String.Empty : currentObject[Guids.Nomenclature.Fields.format].ToString();
            pars.format = format.Contains("*") ? GetSpecRemark(currentObject) : format;
            // Заполнение строк таблицы формы 2
            pars.tableForm2 = GetForm2TableStrings(makes);
            // Получение списков абонентов и заполнение строк таблицы формы 2а
            var externCustomers = GetExternCustomers((ReferenceObject)makes.First());                         // Только для основного исполнения
            var internCustomers = GetInternCustomers((ReferenceObject)makes.First());                         // Только для основного исполнения
            pars.tableForm2a = GetForm2aTableStrings(externCustomers, internCustomers);

            return pars;
        }
        // Получение номера исполнения обьекта номенклатуры
        private int GetMakeNumber(NomenclatureObject currentObject)
        {
            var denotation = currentObject[Guids.Nomenclature.Fields.denotation].GetString();
            var regexResult = RegexPatterns.makeNumber.Matches(denotation);
            return regexResult.Count == 1 ? Convert.ToInt32(regexResult[0].Value) : 0;           // Результат по фрагменту обозначения
        }
        private string GetMakeNumberString(NomenclatureObject currentObject)
        {
            var denotation = currentObject[Guids.Nomenclature.Fields.denotation].GetString();
            var regexResult = RegexPatterns.makeNumber.Matches(denotation);
            var baseVersion = currentObject.GetObject(Guids.Nomenclature.Links.baseMake);
            return regexResult.Count == 1 && baseVersion != null ? regexResult[0].Value.ToString() : String.Empty;       // Результат по фрагменту обозначения
        }
        // Значение примечания спецификации (для тех, у кого формат - составной и указан в примечании)
        private string GetSpecRemark(ReferenceObject currentObject)
        {
            string result = String.Empty;
            // Любой родительский объект
            var parents = currentObject.Parents.Where(t => currentObject[Guids.Nomenclature.Fields.denotation].GetString().Contains(t[Guids.Nomenclature.Fields.denotation].GetString()));
            var parent = parents.Count() > 0 ? parents.FirstOrDefault() : currentObject.Parents.FirstOrDefault();

            if (parent != null)
            {
                // Иерархическая ссылка на currentObject
                var hlink = parent.GetChildLink(currentObject) as NomenclatureHierarchyLink;
                // Примечание спецификации
                result = String.IsNullOrEmpty(hlink.Remarks) ? String.Empty : hlink.Remarks.ToString().Replace("*", "").Replace(")", "");
            }

            return result;
        }
        #region Применяемость объекта и его исполнений
        // Получение строк таблицы лицевой стороны карточки отчета (Форма 2)
        private Dictionary<int, List<string>> GetForm2TableStrings(List<NomenclatureObject> makes)
        {
            var tableStrings = new Dictionary<int, List<string>>();
            // Список применяемостей документа и всех его исполнений
            var useList = new List<DocumentUseString>();
            foreach (var item in makes)
            {
                var make = (ReferenceObject)item;
                useList.AddRange(GetDocumentUse(make));
            }
            useList = useList.OrderBy(t => t.makeNumber).ThenBy(t => t.useLevel).ThenBy(t => t.directUse).ThenBy(t => t.complexUse).ToList();

            // Список изменений документа (только для основного исполнения)
            var modsList = new List<DocumentMod>();
            modsList.AddRange(GetDocumentMods((ReferenceObject)makes.First()));
            // Сколько строк д.б. в таблице
            var stringCount = Math.Max(useList.Sum(t => t.useList.Count == 0 ? 1 : t.useList.Count), 
                modsList.Sum(t => t.contentList.Count == 0 ? 1 : t.contentList.Count));
            // Список из пустых строк
            for (int i = 0; i < stringCount; i++)
            {
                var tableString = new List<string> { String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty };
                tableStrings.Add(i, tableString);
            }
            // Добиваем из списка useList - списка применяемостей всех исполнений данного объекта
            int cnt = 0;
            foreach (var item in useList)
            {
                var currentString = tableStrings[cnt];
                currentString[4] = item.date.Year == 1 ? String.Empty : String.Format(item.date.ToString("dd.MM.yy"));
                if (item.useList.Count > 0)
                {
                    cnt--;
                    foreach (var str in item.useList)
                    {
                        cnt++;
                        currentString = tableStrings[cnt];
                        if (item.useLevel == 0)
                            currentString[5] = str + "firstUsage";
                        else
                            currentString[5] = str;
                    }
                }
                cnt++;
            }
            // Добиваем из списка modsList (предварительно убрав повторяющиеся записи и отсортировав по номеру)
            cnt = 0;
            var distinctModsList = modsList.GroupBy(t => new { t.modNumber, t.denotation, t.date, t.content })
                .Select(g => g.First()).OrderBy(t => t.modNumber).ToList();
            foreach (var mod in distinctModsList)
            {
                var currentString = tableStrings[cnt];
                currentString[6] = mod.modNumber.ToString();
                currentString[7] = mod.denotation;
                currentString[8] = mod.date.Year == 1 ? String.Empty : String.Format(mod.date.ToString("dd.MM.yy"));
                // Обработка содержания изменения (многострочный текст)
                var strs = GetContentList(mod.content, 20f);
                if (strs.Count() > 0)
                {
                    cnt--;
                    foreach (var str in strs)
                    {
                        cnt++;
                        currentString = tableStrings[cnt];
                        currentString[9] = str;
                    }
                }
                cnt++;
            }
            return tableStrings;
        }
        // Получение объекта класса "Использование документа"
        private List<DocumentUseString> GetDocumentUse(ReferenceObject currentObject)
        {
            List<DocumentUseString> useList = new List<DocumentUseString>();
            // Коллекция применяемости currentObject
            var currentObjectUse = new Dictionary<ReferenceObject, List<ReferenceObject>>();
            var parentsList = currentObject.Parents.GroupBy(t => t.SystemFields.Id).Select(t => t.First()).ToList();

            parentsList.ForEach(t => currentObjectUse.Add(t, GetParentObjects(t).GroupBy(g => g.SystemFields.Id)
                .Select(g => g.First()).ToList()));
            // Номер исполнения
            var makeNumber = GetMakeNumberString((NomenclatureObject)currentObject);
            // Первичная применяемость документа
            var firstUseDirect = GetFirstUse(currentObject, true, 0) == null ? String.Empty : 
                GetFirstUse(currentObject, true, 0)[Guids.Nomenclature.Fields.denotation].ToString();
            var firstUseComplex = GetFirstUse(currentObject, false, 0) == null ? String.Empty : 
                GetFirstUse(currentObject, false, 0)[Guids.Nomenclature.Fields.denotation].ToString();

            foreach (var item in currentObjectUse)
            {
                var itemMatches = RegexPatterns.PatternESKD.Matches(item.Key[Guids.Nomenclature.Fields.denotation].ToString());
                foreach (var value in item.Value)
                {
                    var valueMatches = RegexPatterns.PatternESKD.Matches(value[Guids.Nomenclature.Fields.denotation].ToString());
                    if (itemMatches.Count == 0 || valueMatches.Count == 0)
                        continue;

                    var docUseString = new DocumentUseString();

                    var directUseString = itemMatches[0].Value.Trim();
                    var complexUseString = valueMatches[0].Value.Trim();
                    var fullUseString = String.Empty;
                    var useString = String.Empty;
                    if (item.Key == value)
                        useString = directUseString + ";";
                    else
                        useString = directUseString + " -> " + complexUseString + ";";
                    fullUseString = String.IsNullOrEmpty(makeNumber) ? useString : "(" + makeNumber + ") -> " + useString;

                    int cnt = 2;
                    if (item.Key[Guids.Nomenclature.Fields.denotation].ToString() == firstUseDirect)
                    {
                        cnt--;
                        if (value[Guids.Nomenclature.Fields.denotation].ToString() == firstUseComplex)
                            cnt--;
                    }

                    docUseString.makeNumber = String.IsNullOrEmpty(makeNumber) ? 0 : Convert.ToInt32(makeNumber);
                    docUseString.useLevel = cnt;
                    docUseString.date = new DateTime(1, 1, 1);
                    docUseString.directUse = directUseString;
                    docUseString.complexUse = complexUseString;
                    docUseString.useList = GetContentList(fullUseString, 71f).ToList();

                    useList.Add(docUseString);
                }
            }

            return useList.GroupBy(t => new { t.makeNumber, t.useLevel, t.date, t.directUse, t.complexUse }).Select(g => g.First()).ToList();
        }
        // Получение родительских объектов
        private List<ReferenceObject> GetParentObjects(ReferenceObject currentObject)
        {
            var parentObjects = new List<ReferenceObject>();
            if (currentObject.Parents.Count() > 0)
            {
                foreach (var parent in currentObject.Parents)
                    parentObjects.AddRange(GetParentObjects(parent));
            }
            else
            {
                if (currentObject.Class.IsInherit(Guids.Nomenclature.Types.complement) || currentObject.Class.IsInherit(Guids.Nomenclature.Types.complex))
                    parentObjects.Add(currentObject);
            }
            return parentObjects;
        }
        // Получение первичной применяемости объекта
        private ReferenceObject GetFirstUse(ReferenceObject currentObject, bool first, int cnt)
        {
            if (currentObject == null || currentObject.Class.Guid == Guids.Nomenclature.Types.complex || (first && cnt == 1))
                return currentObject;
            else
            {
                string firstUse;
                try
                {
                    firstUse = currentObject[Guids.Nomenclature.Fields.firstUse].ToString();
                }
                catch
                {
                    firstUse = String.Empty;
                }
                if (String.IsNullOrEmpty(firstUse) || firstUse == currentObject[Guids.Nomenclature.Fields.denotation].ToString())
                    return null;
                else
                {
                    var firstUseObject = _nomReference.Find(_nomReference.ParameterGroup[Guids.Nomenclature.Fields.denotation], firstUse).FirstOrDefault();
                    cnt++;
                    return GetFirstUse(firstUseObject, first, cnt);
                }
            }
        }
        #endregion
        #region Изменения объекта и его исполнений
        // Получение объекта класса DocumentMod (изменение объекта)
        private List<DocumentMod> GetDocumentMods(ReferenceObject currentObject)
        {
            var modsList = new List<DocumentMod>();

            var documentMods = currentObject.GetObjects(Guids.Nomenclature.Links.documentMods);
            var currentDenotation = currentObject[Guids.Nomenclature.Fields.denotation].ToString();

            foreach (var item in documentMods)
            {
                var mod = new DocumentMod();
                var currentAction = GetCurrentAction(item, currentDenotation);
                var denotation = item[Guids.DocumentMods.Fields.denotation];
                mod.modNumber = currentAction == null ? 0 : (int)currentAction[Guids.Actions.number];
                mod.denotation = denotation.IsNull ? String.Empty :
                    denotation.ToString();
                mod.date = (DateTime)item[Guids.DocumentMods.Fields.date];
                var content = currentAction == null ? String.Empty : currentAction[Guids.Actions.content].ToString();
                mod.content = content;
                mod.contentList.AddRange(GetContentList(content, 20f));
                modsList.Add(mod);
            }
            return modsList;
        }
        // Получение действия Извещения об изменении
        private ReferenceObject GetCurrentAction(ReferenceObject mod, string denotation)
        {
            var refInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.DocumentMods.id);
            var reference = refInfo.CreateReference();
            reference.LoadSettings.AddRelation(Guids.DocumentMods.Links.actions);
            reference.LoadSettings.AddRelation(Guids.Actions.actionObject);
            reference.LoadSettings.AddParameters(Guids.Nomenclature.Fields.denotation);

            var filter = new Filter(refInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, mod.SystemFields.Id);
            var newMod = reference.Find(filter).First();

            ReferenceObject currentAction = null;

            var actions = newMod.GetObjects(Guids.DocumentMods.Links.actions).ToList();//.Where(t => t[Guids.Actions.number].GetInt32() > 0).ToList();
            foreach (var item in actions)
            {
                var actionObject = item.GetObject(Guids.Actions.actionObject);
                if (actionObject != null && actionObject[Guids.Nomenclature.Fields.denotation].GetString() == denotation)
                {
                    currentAction = item;
                    break;
                }    
            }
            /*var currentAction = actions.Where(t => t.GetObject(Guids.Actions.actionObject)
            [Guids.Nomenclature.Fields.denotation].ToString() == denotation).FirstOrDefault();*/
            return currentAction;
        }
        #endregion
        #region Абоненты и копии объекта и его исполнений
        // Получение строк таблицы обратной стороны карточки отчета (Форма 2а) на основе списков абонентов
        private Dictionary<int, List<string>> GetForm2aTableStrings(List<Customer> externCustomers, List<Customer> internCustomers)
        {
            var tableStrings = new Dictionary<int, List<string>>();

            // Количество строк с учетом того, что в одной строке - 11 ячеек, а на одном листе - 24 строки
            int stringCount = Convert.ToInt32(Math.Ceiling(Math.Max(externCustomers.Count, internCustomers.Count) / 11.0)) * 24;

            // Список из пустых строк
            for (int i = 0; i < stringCount; i++)
            {
                var tableString = new List<string> { String.Empty, String.Empty, String.Empty, String.Empty,
                                                     String.Empty, String.Empty, String.Empty, String.Empty,
                                                     String.Empty, String.Empty, String.Empty };
                tableStrings.Add(i, tableString);
            }

            int cnt = 0;
            int stringCounter = 0;
            // Добиваем из списка externCustomers
            foreach (var item in externCustomers)
            {
                int stringIndex = 24 * stringCounter;
                var currentString = tableStrings[stringIndex];
                currentString[cnt] = item.customer+"Header";
                currentString = tableStrings[stringIndex + 2];
                currentString[cnt] = item.date.Year == 1 ? String.Empty : String.Format(item.date.ToString("dd.MM.yy"));
                currentString = tableStrings[stringIndex + 3];
                currentString[cnt] = item.makeNumber == String.Empty ? item.count.ToString() : "(" + item.makeNumber.Replace("-","") + ") -> " + item.count;
                currentString = tableStrings[stringIndex + 4];
                currentString[cnt] = item.based;
                currentString = tableStrings[stringIndex + 5];
                currentString[cnt] = item.cancelled ? "V" : "";
                cnt++;
                if (cnt == 11)
                {
                    stringCounter++;
                    cnt = 0;
                }
            }
            cnt = 0;
            stringCounter = 0;
            // Добиваем из списка internCustomers
            foreach (var item in internCustomers)
            {
                int stringIndex = 10 + 24 * stringCounter;
                var currentString = tableStrings[stringIndex];
                currentString[cnt] = item.customer+"Header";
                currentString = tableStrings[stringIndex + 2];
                currentString[cnt] = item.date.Year == 1 ? String.Empty : String.Format(item.date.ToString("dd.MM.yy"));
                currentString = tableStrings[stringIndex + 3];
                currentString[cnt] = item.makeNumber == String.Empty ? item.count.ToString() : "(" + item.makeNumber.Replace("-", "") + ") -> " + item.count;
                currentString = tableStrings[stringIndex + 4];
                currentString[cnt] = item.based;
                currentString = tableStrings[stringIndex + 5];
                currentString[cnt] = item.cancelled ? "V" : "";
                cnt++;
                if (cnt == 11)
                {
                    stringCounter++;
                    cnt = 0;
                }
            }
            return tableStrings;
        }
        private List<Customer> GetExternCustomers(ReferenceObject currentObject)
        {
            var customers = new List<Customer>();
            var makeNumber = ((NomenclatureObject)currentObject).Version.IsEmpty ? String.Empty :
                ((NomenclatureObject)currentObject).Version.ToString().Replace("-", "");
            var customerList = currentObject.GetObjects(Guids.Nomenclature.Links.externCustomers);

            foreach (var item in customerList)
            {
                var customer = new Customer();
                customer.makeNumber = makeNumber;
                customer.customer = item[Guids.Nomenclature.ExternCustomer.Fields.customer].IsNull ? String.Empty : item[Guids.Nomenclature.ExternCustomer.Fields.customer].ToString();
                customer.date = (DateTime)item[Guids.Nomenclature.ExternCustomer.Fields.date];
                customer.count = item[Guids.Nomenclature.ExternCustomer.Fields.count].IsNull ? 0 : (int)item[Guids.Nomenclature.ExternCustomer.Fields.count];
                customer.based = item[Guids.Nomenclature.ExternCustomer.Fields.based].IsNull ? String.Empty : item[Guids.Nomenclature.ExternCustomer.Fields.based].ToString();
                customer.cancelled = item[Guids.Nomenclature.ExternCustomer.Fields.cancelled].IsNull ? false : (bool)item[Guids.Nomenclature.ExternCustomer.Fields.cancelled];
                customers.Add(customer);
            }

            return customers;
        }
        private List<Customer> GetInternCustomers(ReferenceObject currentObject)
        {
            var customers = new List<Customer>();
            var makeNumber = ((NomenclatureObject)currentObject).Version.IsEmpty ? String.Empty :
                ((NomenclatureObject)currentObject).Version.ToString().Replace("-", "");
            var customerList = currentObject.GetObjects(Guids.Nomenclature.Links.internCustomers);

            foreach (var item in customerList)
            {
                var customer = new Customer();
                customer.makeNumber = makeNumber;
                customer.customer = item[Guids.Nomenclature.InternCustomer.Fields.customer].IsNull ? String.Empty : item[Guids.Nomenclature.InternCustomer.Fields.customer].ToString();
                customer.date = (DateTime)item[Guids.Nomenclature.InternCustomer.Fields.date];
                customer.count = item[Guids.Nomenclature.InternCustomer.Fields.count].IsNull ? 0 : (int)item[Guids.Nomenclature.InternCustomer.Fields.count];
                customer.based = item[Guids.Nomenclature.InternCustomer.Fields.based].IsNull ? String.Empty : item[Guids.Nomenclature.InternCustomer.Fields.based].ToString();
                customer.cancelled = item[Guids.Nomenclature.InternCustomer.Fields.cancelled].IsNull ? false : (bool)item[Guids.Nomenclature.InternCustomer.Fields.cancelled];
                customers.Add(customer);
            }

            return customers;
        }
        #endregion
        // Разбиение на подстроки многострочного текста
        private string[] GetContentList(string content, double maxWidth)
        {
            return _tForm.Wrap(content, maxWidth);
        }

        public DocumentParameters GetSTODocumentData(ReferenceObject currentObject)
        {
            // Заполнение параметров текущего объекта
            var pars = new DocumentParameters();
            pars.name = currentObject[Guids.STO.Fields.name] == null ? string.Empty : currentObject[Guids.STO.Fields.name].ToString();
            pars.denotation = currentObject[Guids.STO.Fields.denotation] == null ? string.Empty : currentObject[Guids.STO.Fields.denotation].ToString();
            pars.invnum = currentObject[Guids.STO.BTDGroup.Fields.invnum] == null ? string.Empty : currentObject[Guids.STO.BTDGroup.Fields.invnum].ToString();

            var department = currentObject[Guids.STO.BTDGroup.Fields.department] == null ? string.Empty : currentObject[Guids.STO.BTDGroup.Fields.department].ToString();
            var depDigits = RegexPatterns.digits.Matches(department);
            // Подразделение (выделение цифровой части из названия)
            pars.department = depDigits.Count > 0 ? depDigits[0].Value : department;

            pars.sheetsCount = currentObject[Guids.STO.BTDGroup.Fields.sheetsCount].ToString();
            pars.format = "А4";
            pars.doctype = "бумага";
            pars.manufactoryCode = currentObject[Guids.STO.BTDGroup.Fields.manufactoryCode] == null ? string.Empty : currentObject[Guids.STO.BTDGroup.Fields.manufactoryCode].ToString();
            pars.original = pars.manufactoryCode;

            var date = ((DateTime)currentObject[Guids.STO.BTDGroup.Fields.incomingDate]);
            pars.incomingDate = date.Year == 1 ? String.Empty : String.Format(date.ToString("dd.MM.yy"));
            // Заполнение строк таблицы формы 2 (изменения документа)
            pars.tableForm2 = GetForm2TableStringsSTO(currentObject);
            // Получение списков абонентов и заполнение строк таблицы формы 2а
            var externCustomers = GetSTOExternCustomers(currentObject);
            var internCustomers = GetSTOInternCustomers(currentObject);
            pars.tableForm2a = GetForm2aTableStrings(externCustomers, internCustomers);

            return pars;
        }
        private List<Customer> GetSTOInternCustomers(ReferenceObject currentObject)
        {
            var customers = new List<Customer>();
            var customerList = currentObject.GetObjects(Guids.STO.Links.internCustomers);

            foreach (var item in customerList)
            {
                var customer = new Customer();
                customer.customer = item[Guids.STO.InternCustomer.Fields.customer].IsNull ? String.Empty : item[Guids.STO.InternCustomer.Fields.customer].ToString();
                customer.makeNumber = "";
                customer.date = (DateTime)item[Guids.STO.InternCustomer.Fields.date];
                customer.count = item[Guids.STO.InternCustomer.Fields.count].IsNull ? 0 : (int)item[Guids.STO.InternCustomer.Fields.count];
                customer.based = item[Guids.STO.InternCustomer.Fields.based].IsNull ? String.Empty : item[Guids.STO.InternCustomer.Fields.based].ToString();
                customer.cancelled = item[Guids.STO.InternCustomer.Fields.cancelled].IsNull ? false : (bool)item[Guids.STO.InternCustomer.Fields.cancelled];
                customers.Add(customer);
            }

            return customers;
        }
        private List<Customer> GetSTOExternCustomers(ReferenceObject currentObject)
        {
            var customers = new List<Customer>();
            var customerList = currentObject.GetObjects(Guids.STO.Links.externCustomers);

            foreach (var item in customerList)
            {
                var customer = new Customer();
                customer.customer = item[Guids.STO.ExternCustomer.Fields.customer].IsNull ? String.Empty : item[Guids.STO.ExternCustomer.Fields.customer].ToString();
                customer.makeNumber = "";
                customer.date = (DateTime)item[Guids.STO.ExternCustomer.Fields.date];
                customer.count = item[Guids.STO.ExternCustomer.Fields.count].IsNull ? 0 : (int)item[Guids.STO.ExternCustomer.Fields.count];
                customer.based = item[Guids.STO.ExternCustomer.Fields.based].IsNull ? String.Empty : item[Guids.STO.ExternCustomer.Fields.based].ToString();
                customer.cancelled = item[Guids.STO.ExternCustomer.Fields.cancelled].IsNull ? false : (bool)item[Guids.STO.ExternCustomer.Fields.cancelled];
                customers.Add(customer);
            }

            return customers;
        }

        private Dictionary<int, List<string>> GetForm2TableStringsSTO(ReferenceObject currentObject)
        {
            var tableStrings = new Dictionary<int, List<string>>();

            var modsList = GetDocumentModsSTO(currentObject).OrderBy(t => t.content).ThenBy(t => t.modNumber);

            // Список из пустых строк
            for (int i = 0; i < modsList.Count(); i++)
            {
                var tableString = new List<string> { String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty };
                tableStrings.Add(i, tableString);
            }
            int cnt = 0;
            foreach (var mod in modsList)
            {
                var currentString = tableStrings[cnt];
                currentString[6] = mod.modNumber.ToString();
                currentString[7] = mod.denotation;
                currentString[8] = mod.date.Year == 1 ? String.Empty : String.Format(mod.date.ToString("dd.MM.yy"));
                currentString[9] = mod.content;
                cnt++;
            }

            return tableStrings;
        }

        // Получение объекта класса DocumentMod (изменение объекта) документа СТО
        private List<DocumentMod> GetDocumentModsSTO(ReferenceObject currentObject)
        {
            var modsList = new List<DocumentMod>();

            var docFiles = currentObject.GetObjects(Guids.STO.Links.documents);

            foreach (var docFile in docFiles)
            {
                var name = docFile[Guids.Files.Fields.name].GetString();
                var modNumber = docFile[Guids.Files.Fields.modNumber].GetInt32();
                var modDenotation = docFile[Guids.Files.Fields.modDenotation].GetString();
                var part = String.Empty;

                // Выделение фрагмента, связанного с приложением
                var partDocFromName = RegexPatterns.partDoc.Match(name);
                if (partDocFromName.Success)
                    part = partDocFromName.Value;

                if (modNumber == 1 && String.IsNullOrEmpty(modDenotation))
                {
                    // Поиск номера изменения в наименовании файла
                    var modNumberFromName = RegexPatterns.modNumber.Match(name.Replace(" ",""));
                    if (modNumberFromName.Success)
                    {
                        // Поиск обозначения ИИ в наименовании файла
                        var modDoc = RegexPatterns.docPattern.Match(name);
                        if (modDoc.Success)
                        {
                            // Выделение только цифр
                            var modNumberDigits = RegexPatterns.digits.Match(modNumberFromName.Value);
                            if (modNumberDigits.Success)
                            {
                                var mod = new DocumentMod();
                                mod.modNumber = Convert.ToInt32(modNumberDigits.Value);
                                mod.denotation = modDoc.Value;
                                mod.content = part;
                                mod.date = new DateTime(1, 1, 1);
                                modsList.Add(mod);
                            }
                        }
                    }
                }
                else
                {
                    // Заполнение из параметров
                    var mod = new DocumentMod();
                    mod.modNumber = modNumber;
                    mod.denotation = modDenotation;
                    mod.content = part;
                    mod.date = new DateTime(1, 1, 1);
                    modsList.Add(mod);
                }
            }
            return modsList;
        }
    }
}
