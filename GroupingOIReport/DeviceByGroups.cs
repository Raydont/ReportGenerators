using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

namespace GroupingOIReport
{
    public class DeviceByGroups
    {
        public NomenclatureObject Device;
        public double Count;
        public string Number;
        public string NameParents;
        public string Remark = string.Empty;
        public bool Error = false;
        public string NameLoadList = string.Empty;
        //Словарь групп 1 - Наименование группы 2 - списки  объектов для данной группы (Наименование и количество)
        public Dictionary<string, List<NomSpecificationItems>> Groups = new Dictionary<string, List<NomSpecificationItems>>();
        public List<NomSpecificationItems> NotFoundedObject = new List<NomSpecificationItems>();
        static ReferenceInfo nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
        static Reference reference = nomenclatureReferenceInfo.CreateReference();

        public DeviceByGroups()
        {

        }

        public DeviceByGroups(ReportParameters reportParameters, NomSpecificationItems device, LogDelegate logDelegate)
        {
            TextNormalizer textNorm = new TextNormalizer();
            int errorId = 0;
            Device = device.NomObject;
            Count = device.Count;
            Number = device.Number;
            NameLoadList = device.NameLoadList;
            Error = device.Error;

            if (device.NameParents != string.Empty)
            {
                NameParents = device.NameParents + " -> ";
            }

            try
            {
                errorId = 11;
                var hLink = (NomenclatureHierarchyLink) device.NomObject.GetParentHierarchyLinks().FirstOrDefault();

                var children = GetChildren(reportParameters, hLink, null, 1, logDelegate, device.NomObject);
                errorId = 1;
                if (children == null)
                    return;

                foreach (var obj in children)
                {
                    errorId = 2;
                    bool isFindedObject = false;
                    //Получение нормализованных  форм
                    var nameNom = obj.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
                    var denotationNom = obj.NomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
                    var deviceName = device.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
                    var normalNameDenotation = textNorm.GetNormalForm(nameNom + denotationNom);

                    errorId = 3;
                    foreach (var group in GroupsReferenceNames.supplyNormGroup)
                    {
                        errorId = 4;
                        List<NomSpecificationItems> currentGroup = new List<NomSpecificationItems>();

                        //Если наименование позиции начинается с указанной
                        var findedPki = group.Value.FirstOrDefault(t => normalNameDenotation.IndexOf(t.Item2) == 0);

                        // foreach (var findedPki in findedPkis)
                        if (findedPki.Item3 != null && !isFindedObject)
                        {
                            obj.Engineer = findedPki.Item3;
                            errorId = 7;

                            //Замена АИСТ на розетку
                            if ((denotationNom.Contains("АИСТ.469533.001") || denotationNom.Contains("АИСТ.469532.001")))
                            {
                                if (children.Count(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString().Trim().ToLower().Contains("розетка снп260")) == 0)
                                {
                                    var socket = SearchSocketByCell(denotationNom);
                                    if (socket != null)
                                    {
                                        var newNomSpecObj = new NomSpecificationItems(socket, obj.Count, "Добавлено, поскольку отсутствуют розетки. Вместо " + obj.NomObject, findedPki.Item3);
                                        if (newNomSpecObj.NomObject != null)
                                        {
                                            isFindedObject = true;
                                            currentGroup.Add(newNomSpecObj);
                                        }
                                    }
                                    else
                                    {
                                        obj.AddRemark("Ошибка! Не удалось найти и вставить розетки по неопознанному обозначению " + obj.NomObject);
                                        obj.Error = true;
                                        isFindedObject = true;
                                        currentGroup.Add(obj);
                                    }
                                }
                            }
                            //Если находим сборку Катушка индуктивности, соответствующиую регулярному выражению среди потомков, то сердечники не добавляем
                            Regex coilRegex = new Regex(@"Катушка индуктивности\s?(А|Б|В)");
                            if ((nameNom.Trim().ToLower().Contains("сердечник чашечный") ||
                                nameNom.Trim().ToLower().Contains("сердечник подстроечный")) &&
                                children.Count(t => coilRegex.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Success) == 0)
                            {
                            }
                            //Добавление магнита если устройство 1Е036
                            if (denotationNom.Trim().Contains("ГН7.770.008"))
                            {
                                var newNomSpecObj = new NomSpecificationItems(SearchMagnet(), obj.Count * 2, "Вместо " + obj.NomObject + " " + obj.Count + " шт.", findedPki.Item3);
                                if (newNomSpecObj.NomObject != null)
                                {
                                    isFindedObject = true;
                                    currentGroup.Add(newNomSpecObj);
                                }
                            }
                            //Добавление ТБНС, если находим шайбу БФМИ.758491
                            if ((denotationNom.Contains("БФМИ.758491.006") || denotationNom.Contains("БФМИ.758491.054") ||
                                denotationNom.Contains("БФМИ.758491.077") || denotationNom.Contains("БФМИ.758491.076") ||
                                denotationNom.Contains("БФМИ.758491.078") || denotationNom.Contains("БФМИ.758491.079"))
                                && nameNom.Contains("Шайба"))
                            {
                                var newNomSpecObj = new NomSpecificationItems(SearchPreparationByWasher(denotationNom), obj.Count, "Вместо " + obj.NomObject, findedPki.Item3);
                                if (newNomSpecObj.NomObject != null)
                                {
                                    isFindedObject = true;
                                    currentGroup.Add(newNomSpecObj);
                                }
                            }
                            //Добавление шайбы ТКБ6.00.021Ку-02
                            if (denotationNom.Contains("ТКБ6.00.021Ку-02") && nameNom.Contains("Шайба"))
                            {
                                isFindedObject = true;
                                currentGroup.Add(obj);
                            }
                            //Добавление светофильтра
                            if (denotationNom.Contains("203561") && nameNom.Contains("Светофильтр"))
                            {
                                var newNomSpecObj = new NomSpecificationItems(SearchOpticalFilter(), obj.Count, "Вместо " + obj.NomObject, findedPki.Item3);
                                if (newNomSpecObj.NomObject != null)
                                {
                                    isFindedObject = true;
                                    currentGroup.Add(newNomSpecObj);
                                }
                            }
                            //Добавление вилки
                            if (denotationNom.Contains("ГН6.605.035-02") || denotationNom.Contains("5В55.7205.0-5"))
                            {
                                var newNomSpecObj = new NomSpecificationItems(SearchFork(), 8, "Вместо " + obj.NomObject, findedPki.Item3);
                                if (newNomSpecObj.NomObject != null)
                                {
                                    isFindedObject = true;
                                    currentGroup.Add(newNomSpecObj);
                                }
                            }
                            //Добавление объекта в группу
                            currentGroup.Add(obj);
                            isFindedObject = true;
                        }

                        errorId = 9;
                        //Если группа уже есть в списке то добавлем объекты к ней, иначе создаем такую группы и записываем объекты
                        if (Groups.Keys.Contains(group.Key))
                        {
                            for (int i = 0; i < Groups.Count; i++)
                            {
                                if (Groups.Keys.ToList()[i] == group.Key)
                                {
                                    foreach (var groupObject in currentGroup)
                                    {
                                        Groups.Values.ToList()[i].Add(groupObject);
                                    }
                                }
                            }
                        }
                        else
                        {
                            Groups.Add(group.Key, currentGroup);
                        }
                    }
                    if (!isFindedObject)
                    {
                        if (NotFoundedObject.Count(t => t.NomObject == obj.NomObject) > 0)
                        {
                            NotFoundedObject.Where(t => t.NomObject == obj.NomObject).FirstOrDefault().Count += obj.Count;
                            if (obj.Remarks.Count > 0)
                            {
                                NotFoundedObject.Where(t => t.NomObject == obj.NomObject).FirstOrDefault().AddRemarks(obj.Remarks);
                            }
                        }
                        else
                        {
                            NotFoundedObject.Add(obj);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Обратитесь в отдел 911. Сформированный отчет будет некорректен. Код ошибки 5. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки " + errorId + "\r\n" + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        //Поиск и вставка розетки, по обозначению ячейки АИСТ.469533.001, АИСТ.469532.001
        public NomenclatureObject SearchSocketByCell(string denotation)
        {
            Regex regexMakeCell = new Regex(@"(?<=(АИСТ\.469532\.001-|АИСТ\.469533\.001-))\d{1,}");
            string make = regexMakeCell.Match(denotation).Value.ToString();

            var idSocket = make switch
            {
                //Розетка Р.СНП34-135/132=1                
                "001" => 860,
                "05" => 860,
                "10" => 860,
                "15" => 860,
                "20" => 860,
                "25" => 860,
                "30" => 860,
                "35" => 860,
                //Розетка Р.СНП34-113/132=1
                "01" => 13043,
                "06" => 13043,
                "11" => 13043,
                "16" => 13043,
                "021"=> 13043,
                "26" => 13043,
                "31" => 13043,
                "36" => 13043,
                //Розетка Р.СНП34-90/132=1
                "02" => 427,
                "07" => 427,
                "12" => 427,
                "17" => 427,
                "22" => 427,
                "27" => 427,
                "32" => 427,
                "37" => 427,
                //Розетка Р.СНП34-69/132=1
                "03" => 243781,
                "08" => 243781,
                "13" => 243781,
                "18" => 243781,
                "23" => 243781,
                "28" => 243781,
                "33" => 243781,
                "38" => 243781,
                //Розетка Р.СНП34-46/132=1
                "04" => 1020,
                "09" => 1020,
                "14" => 1020,
                "19" => 1020,
                "24" => 1020,
                "29" => 1020,
                "34" => 1020,
                "39" => 1020,
                _    => 0
            };

            return idSocket == 0 ? null : (NomenclatureObject)reference.Find(idSocket);
        }


        //Поиск магнитов
        public NomenclatureObject SearchMagnet()
        {
            return (NomenclatureObject)reference.Find(4538);
        }

        //Поиск светофильтра
        public NomenclatureObject SearchOpticalFilter()
        {
            return (NomenclatureObject)reference.Find(46836);
        }

        //Поиск вилки
        public NomenclatureObject SearchFork()
        {
            return (NomenclatureObject)reference.Find(9077);
        }

        //Поиск заготовки по шайбе
        public NomenclatureObject SearchPreparationByWasher(string denotation)
        {
            Regex regexMakeCell = new Regex(@"(?<=(БФМИ\.758491\.|БФМИ\.758491\.|БФМИ\.758491\.|БФМИ\.758491\.|БФМИ\.758491\.))\d{3}(-\d{1,})?");
            string make = regexMakeCell.Match(denotation).Value.ToString();
            int idPreparation = 0;
            /*
5573	Заготовка -2.ТБНС, ЫРО.737.001ТУ
5796	Заготовка -8.ТБНС, ЫРО.737.001ТУ
222554	Заготовка 3, ТБНС, ЫР0.737.001 ТУ
222555	Заготовка 1, ТБНС, ЫР0.737.001 ТУ
415083	Заготовка ТБНС11
415098	Заготовка ТБНС5
416537	Заготовка ТБНС-3
416538	Заготовка ТБНС-1
416539	Заготовка ТБНС-16
421156	Заготовка ТБНС-9
421157	Заготовка ТБНС-13 */

            idPreparation = make switch
            {
                //Розетка Р.СНП34-135/132=1
                "006"    => 415098,
                "006-01" => 415083,
                "006-02" => 421156,
                "006-03" => 415083,
                "006-04" => 5796,
                "006-05" => 415083,
                "006-06" => 415083,
                "006-07" => 415083,
                "006-08" => 415098,
                "006-09" => 415098,
                "054"    => 415083,
                "054-01" => 415083,
                "054-02" => 415083,
                "054-03" => 421157,
                "054-04" => 415083,
                "077"    => 416537,
                "077-01" => 5573,
                "076"    => 416539,
                "076-01" => 415083,
                "078"    => 415083
            };
            return (NomenclatureObject)reference.Find(idPreparation);
        }

    //Получение всех дочерних объектов 
        public List<NomSpecificationItems> GetChildren(ReportParameters reportParameters, NomenclatureHierarchyLink childObjectNHL, NomenclatureObject parentObject, double countParent, LogDelegate logDelegate, NomenclatureObject childObject)
        {
            var listObjects = new List<NomSpecificationItems>();
            int errorID = 0;
            try
            {

                //Если устройство шифрованное, и это не родительсктй объект и он  не отмечен флажок "Раскрывать состав шифрованного устройства" возваращаем нуль
                if (reportParameters.NomObjects.Count(t=>t.NomObject == childObject) > 0 && parentObject != null && 
                    !(reportParameters.NomObjects.Any(t => t.NomObject == childObject) && !reportParameters.AddAllCodeDevice && !reportParameters.NotCodeDevices && !reportParameters.AddCodeDevices))// (childObject.Class.IsInherit(Guids.NomenclatureTypes.Assembly) &&
                  //  childObject[Guids.NomenclatureParameters.CodeGuid].Value.ToString().Trim() != string.Empty) && parentObject != null)
                {
                    return null;
                }

                Regex coilRegex = new Regex(@"Катушка индуктивности\s?(А|Б|В)");
                var childName = childObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
                var childDenotation = childObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();


                if ((childObject.Class.IsInherit(Guids.NomenclatureTypes.StandartItem) && reportParameters.StandartItemsType) ||
                    (childObject.Class.IsInherit(Guids.NomenclatureTypes.OtherItem) && reportParameters.OtherItems) ||
                    (childObject.Class.IsInherit(Guids.NomenclatureTypes.Detail) && reportParameters.DetailsItems) ||
                    (childObject.Class.IsInherit(Guids.NomenclatureTypes.Material) && childName.ToLower().Contains("тбнс")) ||
                    (childObject.Class.IsInherit(Guids.NomenclatureTypes.Assembly) && (coilRegex.Match(childName).Success ||
                    (childDenotation.Contains("203561") || childName.Contains("Светофильтр")) ||
                    childDenotation.Contains("5В557260.300-04") ||
                    childDenotation.Contains("2А40Е.16.01.120") ||
                    childDenotation.Contains("2А40Е.16.01.110") ||
                    childDenotation.Contains("5В55.7205.0") ||
                    childDenotation.Contains("ГН6.605.035-02") ||
                    childDenotation.Contains("5В55.7205.0-5") ||
                    GroupsReferenceNames.assemblyList.Count(t => childDenotation.Trim().ToLower().Contains(t.Key.Trim().ToLower())) > 0)) ||
                    (childObject.Class.IsInherit(Guids.NomenclatureTypes.Detail) && (childDenotation.Contains("ГН7.770.008") ||
                    (((childDenotation.Contains("БФМИ.758491.006") || childDenotation.Contains("БФМИ.758491.054") ||
                       childDenotation.Contains("БФМИ.758491.077") || childDenotation.Contains("БФМИ.758491.076") ||
                       childDenotation.Contains("БФМИ.758491.078")) || childDenotation.Contains("ТКБ6.00.021Ку-02") ||
                       childDenotation.Contains("ТКБ6") ||
                       childDenotation.Contains("БФМИ.755415.001")) && childName.Contains("Шайба")))))
                {                  
                    errorID = 2;

                    if (parentObject == null)
                    {
                        return null;
                    }

                    double amount = Convert.ToDouble(childObjectNHL[NomenclatureHierarchyLink.FieldKeys.Amount].Value);

                    var remark = childObjectNHL[NomenclatureHierarchyLink.FieldKeys.Remarks].Value.ToString();
                    if (remark.ToLower().Contains("из состава 694700м") ||
                        remark.ToLower().Contains("из состава вд6.096.012") ||
                        (remark.ToLower().Contains("из состава пэвм \"рамэк\"") && childName.Contains("Вилка")))
                    {
                    }
                    else
                    {
                        if ((amount != Math.Truncate(amount) || childObject.Class.IsInherit(Guids.NomenclatureTypes.OtherItem)) && 
                            (childObject[Guids.NomenclatureParameters.NameGuid].GetString().ToLower().Contains("резист") ||
                            childObject[Guids.NomenclatureParameters.NameGuid].GetString().ToLower().Contains("конден")))
                        {
                            var countCalcRez = CountSelectionRezistor(childObject, remark);
          
                            if (countCalcRez > amount || amount != Math.Truncate(amount))
                            {
                                amount = countCalcRez * countParent;
                            }
                            else
                            {
                                amount = amount * countParent;
                            }
                        }
                        else
                        {
                            amount = amount * countParent;
                        }

                        if (listObjects.Count(t => t.NomObject == childObject) == 0)
                        {
                            errorID = 5;
                            var newObj = new NomSpecificationItems(childObject, amount, remark);
                            newObj.NameParents += parentObject.ToString() + " ";
                            listObjects.Add(newObj);
                        }
                        else
                        {
                            listObjects.Where(t => t.NomObject == childObject).FirstOrDefault().Count += amount;
                         //   listObjects.Where(t => t.NomObject == childObject).FirstOrDefault().NameParents += parentObject.ToString() + " "; 

                            if (amount == 0)
                            {
                                listObjects.Where(t => t.NomObject == childObject).FirstOrDefault().CountError = true;
                            }
                            errorID = 8;
                            if (remark != string.Empty)
                                listObjects.Where(t => t.NomObject == childObject).FirstOrDefault().AddRemark(remark);
                        }
                    }

                }
                else
                {
                    foreach (NomenclatureHierarchyLink childHlink in childObject.Children.GetHierarchyLinks())
                    {
                        double amount = 1;

                        if (parentObject != null)
                        {
                            amount = Convert.ToDouble(childObjectNHL[NomenclatureHierarchyLink.FieldKeys.Amount].Value);
                            var remark = childObjectNHL[NomenclatureHierarchyLink.FieldKeys.Remarks].Value.ToString();                         
                        }

                        //Иерархия номенклатуры
                        errorID = 33;

                        List<NomSpecificationItems> children = new List<NomSpecificationItems>();

                        //Добавляем всех потомков, если параметр первый уровень не выделен, иначе добавляем только потомков на 1 шаге раскрытия
                        if ((reportParameters.NotCodeDevices &&
                            (childObject[Guids.NomenclatureParameters.CodeGuid].Value.ToString().Trim() == string.Empty || parentObject == null)) ||
                            !reportParameters.NotCodeDevices)
                        {

                            amount = amount * countParent;
                            var childrenObjects = GetChildren(reportParameters, childHlink, childObject, amount, logDelegate, (NomenclatureObject)childHlink.ChildObject);
                            if (childrenObjects != null)
                                children.AddRange(childrenObjects);
                        }


                        if (children == null)
                        {
                            continue;
                        }
                        foreach (var childParent in children)
                        {
                            errorID = 12;

                            if (listObjects.Count(t => t.NomObject == childParent.NomObject) == 0)
                            {
                                listObjects.Add(childParent);
                                errorID = 16;
                            }
                            else
                            {
                                errorID = 17;
                                listObjects.Where(t => t.NomObject == childParent.NomObject).FirstOrDefault().Count += childParent.Count;
                             //   listObjects.Where(t => t.NomObject == childParent.NomObject).FirstOrDefault().NameParents += childObject.ToString() + " ";

                                if (childParent.Remarks.Count > 0)
                                    listObjects.Where(t => t.NomObject == childParent.NomObject).FirstOrDefault().AddRemarks(childParent.Remarks);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Внимание ошибка! Обратитесь в отдел 911. Сформированный отчет будет некорректен. Код ошибки 6. " + errorID.ToString()  + " par - " + parentObject + " child - " + childObject + "\r\n" + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return listObjects;
        }

        private int CountSelectionRezistor(NomenclatureObject nomObj, string remark)
        {
            var remarkWithOutLetter = remark.Replace("*", "");
            int errorId = 0;
            Regex regexDigitalOne = new Regex(@"(C|С|R|W)\d{1,}");
            Regex regexDigital = new Regex(@"(?<=(C|С|R|W))\d{1,}");
            Regex regexDigitalTwo = new Regex(@"(C|С|R|W)\d{1,}(-|(\s{1,})?\.{3}(\s{1,})?)(C|С|R|W)\d{1,}");
            int countRezist = 0;
            var matchrez = string.Empty;
            try
            {
                errorId = 1;

                foreach (Match match in regexDigitalTwo.Matches(remarkWithOutLetter))
                {
                    errorId = 2;
                    matchrez = regexDigital.Matches(match.Value.ToString())[0].Value +" " + regexDigital.Matches(match.Value.ToString())[1].Value;
                    countRezist += Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[1].Value) - Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[0].Value) + 1;
                    errorId = 3;
                    remarkWithOutLetter = remarkWithOutLetter.Replace(match.Value.ToString(), "");
                    errorId = 4;
                }
                errorId = 5;
                countRezist += regexDigitalOne.Matches(remarkWithOutLetter).Count;
                errorId = 6;
            }
            catch
            {
                MessageBox.Show("Обратитесь в отдел 911. Сформированный отчет будет некорректен. Код ошибки 7. Не могу посчитать количество резистора\r\n" + nomObj + "\r\nОшибка в примечании: " + remark + "\r\nКод ошибки " + errorId +"\r\n" + matchrez, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
             
            }
            return countRezist;
        }
    }

    public static class GroupsReferenceNames
    {
        public static Dictionary<string, List<(string, string, string)>> supplyNormGroup = new Dictionary<string, List<(string, string, string)>>();

        // Группы снабжения по каждому экономисту. У экономиста список ПКИ с инженером
        public static Dictionary<string, List<KeyValuePair<string, string>>> supplyGroup = new Dictionary<string, List<KeyValuePair<string, string>>>(); 

        public static List<KeyValuePair<string, string>> assemblyList = new List<KeyValuePair<string, string>>();
        public static void Initializer()
        {
            supplyNormGroup = new Dictionary<string, List<(string, string, string)>>();
            supplyGroup = new Dictionary<string, List<KeyValuePair<string, string>>>();
            assemblyList = new List<KeyValuePair<string, string>>();
            //Создаем ссылку на справочник
            ReferenceInfo info = ReferenceCatalog.FindReference(Guids.References.GroupOtherAndStdItem);
            Reference reference = info.CreateReference();
            //Находим тип «Папка»
            ClassObject classObject = info.Classes.Find(Guids.GroupOtherAndStdItemsTypes.Folder);
            //Создаем фильтр
            Filter filter = new Filter(info);
            //Добавляем условие поиска – «Тип = Папка»
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);
            filter.Terms.AddTerm("[Родительский объект].[Guid]", ComparisonOperator.Equal, new Guid("6d40876e-e1c0-452f-9c39-b93403f1b977"));
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр            
            List<ReferenceObject> listObj = reference.Find(filter);

            string objToStr = string.Empty;

            foreach (var folder in listObj.OrderBy(t=>t.ToString()))
            {
                supplyGroup.Add(folder.ToString(),
                    folder.Children.Select(t => new KeyValuePair<string, string> (t.ToString(), t[Guids.GroupOtherAndStdItemsParameters.Engineer].ToString().Trim())).ToList());
            }

            foreach (var folder in listObj)
            {
                assemblyList.AddRange(folder.Children.Where(t => Convert.ToBoolean(t[Guids.GroupOtherAndStdItemsParameters.Assembly].Value) == true)
                    .Select(t => new KeyValuePair<string, string> (t.ToString(), t[Guids.GroupOtherAndStdItemsParameters.Engineer].ToString())));
            }

            TextNormalizer textNorm = new TextNormalizer();

            foreach (var group in supplyGroup)
            {
                var supply = new List<(string, string, string)>();
                foreach (var name in group.Value)
                {
                    supply.Add((name.Key, textNorm.GetNormalForm(name.Key), name.Value));
                }
                supplyNormGroup.Add(group.Key, supply);
            }
        }
    }
    public class NomSpecificationItems
    {
        public NomenclatureObject NomObject;
        public string Engineer = string.Empty;
        public double Count;
        public List<string> Remarks = new List<string>();
        public string Number;
        public bool CountError = false;
        public string NameParents = string.Empty;
        public bool Error = false;

        public string NameLoadList = string.Empty;

        public string Type = string.Empty;
        private string letter = string.Empty;

        public string Letter
        {
            get
            {
                return letter;
            }

            set
            {
                letter = value.Trim().ToLower() switch
                {
                    "c" => "С",
                    "с" => "С",
                    _ => string.Empty
                };
            }
        }

        public NomSpecificationItems (NomenclatureObject nomObject, double count, List<string> remarks, string engineer)
        {
            NomObject = nomObject;
            Count = count;
            Engineer = engineer;
            foreach (var remark in remarks)
            {
                if (!Remarks.Contains(remark) && remark.Trim() != string.Empty)
                    Remarks.Add(remark);
            }
        }

        public NomSpecificationItems(NomenclatureObject nomObject, string nameLoadList, double count)
        {
            NomObject = nomObject;
            NameLoadList = nameLoadList;
            Count = count;
        }


        //public NomSpecificationItems(NomenclatureObject nomObject, double count, string remark, string engineer)
        //{
        //    NomObject = nomObject;
        //    Engineer = engineer;
        //    Count = count;
        //    if (!Remarks.Contains(remark) && remark.Trim() != string.Empty)
        //        Remarks.Add(remark);
        //}

        public NomSpecificationItems(NomenclatureObject nomObject, double count, string remark, string engineer = "", string number = null)
        {
            NomObject = nomObject;
            Engineer = engineer;
            Count = count;      

            if (!Remarks.Contains(remark) && remark.Trim() != string.Empty)
                Remarks.Add(remark);

            if (number != null && number.Trim() != string.Empty)
                Number = number;
        }

        public void AddRemarks(List<string> remarks)
        {
            foreach (var remark in remarks)
            {
                AddRemark(remark);
            }
        }

        public void AddRemark(string remark)
        {
            if (!Remarks.Contains(remark) && remark.Trim() != string.Empty)
                Remarks.Add(remark);
        }
    }
}
