using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Permissions;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Catalogs;
using TFlex.DOCs.Model.References.Nomenclature.ModificationNotices;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Structure;

namespace TPPDataGenerator
{
    public class DataReader
    {
        private readonly MacroProvider macro;
        public DataReader(MacroProvider macro)
        {
            this.macro = macro;
        }
        // Все объекты извещения либо все объекты состава нового КД, не входящие в стоп-лист
        public List<ReferenceObject> ПолучитьПроверяемыеОбъекты(List<ReferenceObject> объектыНоменклатуры)
        {
            var проверяемыеОбъекты = new List<ReferenceObject>();
            var стопЛист = macro.НайтиОбъекты(Guids.СтопЛистДляТПП.id.ToString(), "[ID] > 0").ToList();
            foreach (var объектНоменклатуры in объектыНоменклатуры)
            {
                var наименование = объектНоменклатуры[Guids.Номенклатура.Поля.наименование].GetString();
                var обозначение = объектНоменклатуры[Guids.Номенклатура.Поля.обозначение].GetString();
                if (ВходитВСтопЛисты(стопЛист, наименование, обозначение))
                    continue;
                проверяемыеОбъекты.Add(объектНоменклатуры);
            }
            return проверяемыеОбъекты;
        }
        // ID объектов номенклатуры с неактуальными ТП
        public List<int> ПолучитьIDОбъектовСНеактуальнымиТП(List<ReferenceObject> объектыНоменклатуры, Reference номенклатураСправочник)
        {
            var объектыСНеактуальнымиТП = new List<int>();
            var сдвиг = 0;
            while (сдвиг <= объектыНоменклатуры.Count)
            {
                var объектыНоменклатурыПорция = объектыНоменклатуры.Skip(сдвиг).Take(Settings.сдвиг).ToList();
                if (объектыНоменклатурыПорция.Count > 0)
                {
                    номенклатураСправочник.LoadLinks(объектыНоменклатурыПорция, номенклатураСправочник.LoadSettings);
                    foreach (var объектНоменклатуры in объектыНоменклатурыПорция)
                    {
                        var тпОбъекта = объектНоменклатуры.GetObjects(Guids.Номенклатура.Связи.техПроцессы).ToList();
                        var утвержденныеТП = тпОбъекта.Where(t => t.SystemFields.Stage != null && t.SystemFields.Stage.Guid == Guids.Стадии.ТП.тпУтвержден).ToList();
                        if (тпОбъекта.Count == 0 || утвержденныеТП.Count == 0 || объектНоменклатуры[Guids.Номенклатура.Поля.тпНеактуален].GetBoolean() == true)
                            объектыСНеактуальнымиТП.Add(объектНоменклатуры.SystemFields.Id);
                    }
                }
                сдвиг += Settings.сдвиг;
            }
            return объектыСНеактуальнымиТП;
        }
        public List<СтрокаОтчета> ПолучитьДанныеДляОтчета(List<ReferenceObject> объектыНоменклатуры)
        {
            // В запросе поиск ведется внутри актуальных ЭСЗ по списку ЭСИ, заданных в списке объектыНоменклатуры.
            // Для ИИ или новой КД.
            var справочникЭСЗ = ServerGateway.Connection.ReferenceCatalog.Find(Guids.СтруктурыЗаказов.id);
            var справочникНоменклатура = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Номенклатура.id);

            var объектыДляОтчета_ID = объектыНоменклатуры.Select(t => t.SystemFields.Id).ToList();
            var родительскиеТипыЭСЗ_ID = new List<int> { справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.структура).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.дсеДляРазработки).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.описание).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.вариантЭСИ).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.указаниеДоработка).Id };
            var запрещенныеТипыЭСЗ_ID = new List<int> { справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.документ).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.дсеДляРазработки).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.прочееИзделие).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.стандартноеИзделие).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.описание).Id };
            var строкаЗапроса = string.Format(Settings.Запросы.объектыЭСИвЭСЗ,
                                              string.Join(",", родительскиеТипыЭСЗ_ID),
                                              string.Join(",", запрещенныеТипыЭСЗ_ID),
                                              справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.вариантЭСИ).Id,
                                              string.Join(",", объектыДляОтчета_ID),
                                              справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.документ).Id,
                                              справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.деталь).Id);
            return ЧтениеДанныхИзЭСЗ(строкаЗапроса);
        }
        public List<СтрокаОтчета> ПолучитьДанныеДляОтчетаИзВыгружаемыхЭСЗ(int объектСЗ_ID = 0)
        {
            // В запросе, если объектСЗ_ID = 0, то поиск ведется по списку ЭСЗ, у которых установлен флаг "Выгрузить в 1С",
            // иначе поиск ведется по ЭСЗ, ID которой задан в объектСЗ_ID.
            // Для ЭСЗ.
            var справочникЭСЗ = ServerGateway.Connection.ReferenceCatalog.Find(Guids.СтруктурыЗаказов.id);
            var справочникНоменклатура = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Номенклатура.id);

            var родительскиеТипыЭСЗ_ID = new List<int> { справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.структура).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.дсеДляРазработки).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.описание).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.вариантЭСИ).Id,
                                                         справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.указаниеДоработка).Id };
            var запрещенныеТипыЭСЗ_ID = new List<int> { справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.документ).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.дсеДляРазработки).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.прочееИзделие).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.стандартноеИзделие).Id,
                                                        справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.описание).Id };
            var разрешенныеТипыЭСИ_ID = new List<int> { справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.деталь).Id,
                                                        справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.сборочнаяЕдиница).Id,
                                                        справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.комплект).Id,
                                                        справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.указаниеДоработка).Id };
            var строкаПодзапроса = объектСЗ_ID == 0 ? Settings.Запросы.выгружаемыеЭСЗ : объектСЗ_ID.ToString();
            var строкаЗапроса = string.Format(Settings.Запросы.объектыЭСИвВыгружаемыхЭСЗ,
                                              string.Join(",", родительскиеТипыЭСЗ_ID),
                                              string.Join(",", запрещенныеТипыЭСЗ_ID),
                                              справочникЭСЗ.Classes.Find(Guids.СтруктурыЗаказов.Типы.вариантЭСИ).Id,
                                              string.Join(",", разрешенныеТипыЭСИ_ID),
                                              справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.документ).Id,
                                              справочникНоменклатура.Classes.Find(Guids.Номенклатура.Типы.деталь).Id,
                                              Stage.Find(ServerGateway.Connection, Guids.Стадии.ТП.тпУтвержден).Id)
                                              .Replace("##ЭСЗ##", строкаПодзапроса);
            return ЧтениеДанныхИзЭСЗ(строкаЗапроса);
        }
        private List<СтрокаОтчета> ЧтениеДанныхИзЭСЗ(string строкаЗапроса)
        {
            var строкиОтчета = new List<СтрокаОтчета>();
            var строкиОтчетаПромежуточные = new List<СтрокаОтчета>();
            var стадияТПУтвержден = Stage.Find(ServerGateway.Connection, Guids.Стадии.ТП.тпУтвержден);
            var стопЛист = macro.НайтиОбъекты(Guids.СтопЛистДляТПП.id.ToString(), "[ID] > 0").ToList();
            try
            {
                using (SqlConnection conn = GetConnection(true))
                {
                    SqlCommand команда = new SqlCommand(строкаЗапроса, conn);
                    команда.CommandTimeout = 0;
                    SqlDataReader reader = команда.ExecuteReader();

                    while (reader.Read())
                    {
                        var строкаОтчета = new СтрокаОтчета();
                        строкаОтчета.объектЭСЗ_id = Convert.ToInt32(reader[0]);
                        строкаОтчета.объектЭСЗ_номерЗаказа = Convert.ToString(reader[1]);
                        строкаОтчета.объектЭСЗ_обозначение = Convert.ToString(reader[2]);
                        строкаОтчета.объектЭСИ_id = Convert.ToInt32(reader[3]);
                        строкаОтчета.объектЭСИ_обозначение = Convert.ToString(reader[4]);
                        строкаОтчета.объектЭСИ_наименование = Convert.ToString(reader[5]);
                        строкаОтчета.объектЭСИ_тип = Convert.ToString(reader[6]);
                        строкаОтчета.объектЭСИ_стадияТП = string.IsNullOrEmpty(Convert.ToString(reader[7])) ? "ТП отсутствует" : Convert.ToString(reader[7]);
                        строкаОтчета.объектЭСИ_извещение = string.IsNullOrEmpty(Convert.ToString(reader[8]))
                            ? "" : Convert.ToString(reader[8]) + " от " + Convert.ToString(reader[9]);
                        строкаОтчета.объектЭСИ_примечаниеТП = Convert.ToString(reader[10]);
                        строкаОтчета.объектЭСИ_естьКД = Convert.ToString(reader[11]);
                        // Если входит в стоп-листы - запись объекта пропускается
                        if (ВходитВСтопЛисты(стопЛист, строкаОтчета.объектЭСИ_наименование, строкаОтчета.объектЭСИ_обозначение))
                            continue;
                        строкиОтчетаПромежуточные.Add(строкаОтчета);
                    }
                    var строкиОтчетаЭСЗ_ID = строкиОтчетаПромежуточные.Select(t => t.объектЭСЗ_id).Distinct().ToList();
                    // Если ТП несколько, и среди них хотя бы один - утвержденный, то в отчет пишется он, если же нет - то первый в списке
                    foreach (var строкаЭСЗ_ID in строкиОтчетаЭСЗ_ID)
                    {
                        var строкиЭСИ_ID = строкиОтчетаПромежуточные.Where(t => t.объектЭСЗ_id == строкаЭСЗ_ID).Select(t => t.объектЭСИ_id).Distinct().ToList();
                        foreach (var строкаЭСИ_ID in строкиЭСИ_ID)
                        {
                            var строкиЭСИ = строкиОтчетаПромежуточные.Where(t => t.объектЭСЗ_id == строкаЭСЗ_ID && t.объектЭСИ_id == строкаЭСИ_ID).ToList();
                            if (строкиЭСИ.Where(t => t.объектЭСИ_стадияТП == стадияТПУтвержден.Name).Count() > 0)
                                строкиОтчета.Add(строкиЭСИ.Where(t => t.объектЭСИ_стадияТП == стадияТПУтвержден.Name).First());
                            else
                                строкиОтчета.Add(строкиЭСИ.First());
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                macro.Сообщение("Ошибка!", exc.Message);
            }
            return строкиОтчета;
        }
        // Проверка, не входит ли объект номенклатуры в список ограничений (для которых ТП не разрабатываются)
        private bool ВходитВСтопЛисты(Объекты стопЛист, string наименование, string обозначение)
        {
            foreach (var элемент in стопЛист)
            {
                var элементНаименование = элемент[Guids.СтопЛистДляТПП.Поля.наименование.ToString()].ToString();
                var элементОбозначение = элемент[Guids.СтопЛистДляТПП.Поля.обозначение.ToString()].ToString();
                // Ограничение по наименованию
                if (string.IsNullOrEmpty(элементОбозначение) && наименование.Contains(элементНаименование))
                    return true;
                // Ограничение по обозначению
                if (string.IsNullOrEmpty(элементНаименование) && обозначение.Contains(элементОбозначение))
                    return true;
                // Ограничение по наименованию и обозначению
                if (наименование.Contains(элементНаименование) && обозначение.Contains(элементОбозначение))
                    return true;
            }
            // Детали или сборки с обозначениями документов
            if (Regex.Match(обозначение, Settings.всеДокументыRegex).Success)
                return true;
            // Детали с обозначениями ЕСПД
            if (Regex.Match(обозначение, Settings.ЕСПДRegex).Success)
                return true;
            return false;
        }
        private SqlConnection GetConnection(bool open)
        {
            var строкаПодключения = macro.НайтиОбъект("Глобальные параметры", "[GUID] = " + Guids.строкаПодключения).ToString();
            string connectionString = Decoder.Decrypt(macro.ГлобальныйПараметр[строкаПодключения].ToString());
            try
            {
                //подключение к БД T-FLEX DOCs
                SqlConnection connection = new SqlConnection(connectionString);
                SqlConnection.ClearPool(connection);
                if (open)
                    connection.Open();

                return connection;
            }
            catch (Exception exc)
            {
               macro.Сообщение("Ошибка!", exc.Message);
            }
            return null;
        }
    }
    public static class Decoder
    {
        public static string GetKey()
        {
            var info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            var reference = info.CreateReference();
            return reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
        }
        private static byte[] GenerateValidKey(int length)
        {
            byte[] newBytes = new byte[length];
            var keyBytes = Encoding.UTF8.GetBytes(GetKey());
            Array.Copy(keyBytes, newBytes, Math.Min(length, keyBytes.Length));
            return newBytes;
        }
        public static string Decrypt(string encodedValue)
        {
            byte[] cipherText = Convert.FromBase64String(encodedValue);
            using (Aes aes = Aes.Create())
            {
                aes.Key = GenerateValidKey(32);
                aes.IV = new byte[16]; // Формирование вектора инициализации
                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);
                using (MemoryStream msDecrypt = new MemoryStream(cipherText))
                {
                    using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                        {
                            return srDecrypt.ReadToEnd();
                        }
                    }
                }
            }
        }
    }
}
