using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;

namespace NomenclatureReport
{
    // Тип элемента
    public enum ItemType
    {
        TYPE_COMPLEX = 0,
        TYPE_DETAIL = 1,
        TYPE_ASSEMBLY = 2,
        TYPE_PKI = 3
    }

    // Информация об элементе
    public class ItemInfo
    {
        int _id;                    // ID элемента
        int _parentId;              // ID родительского элемента
        int _level;                 // Уровень вложенности элемента
        string _name;               // Наименование элемента 
        string _denotation;         // Обозначение элемента
        double _amount;              // Количество элемента
        string _unitCode;           // Шифр элемента
        string _marshrut;           // Маршрут элемента
        string _material;           // Материал элемента
        string _bomSection;         // Раздел спецификации
        int _classId;               // ID класса
        int unitCalculated;         // Флаг обсчета элемента

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public int ParentId
        {
            get { return _parentId; }
            set { _parentId = value; }
        }

        public int Level
        {
            get { return _level; }
            set { _level = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string Denotation
        {
            get { return _denotation; }
            set { _denotation = value; }
        }

        public double Amount
        {
            get { return _amount; }
            set { _amount = value; }
        }

        public string UnitCode
        {
            get { return _unitCode; }
            set { _unitCode = value; }
        }

        public string Marshrut
        {
            get { return _marshrut; }
            set { _marshrut = value; }
        }

        public string Material
        {
            get { return _material; }
            set { _material = value; }
        }

        public string BomSection
        {
            get { return _bomSection; }
            set { _bomSection = value; }
        }

        public int ClassId
        {
            get { return _classId; }
            set { _classId = value; }
        }

        public int UnitCalculated
        {
            get { return unitCalculated; }
            set { unitCalculated = value; }
        }
    }

    // Информация о комплексе
    public class ComplexInfo
    {
        int _productId;       // ID в таблице Products (добавленных заказов)
        int _objectId;        // ObjectID заказа в DOCs
        string _unitCode;     // Шифр
        string _name;         // Наименование заказа
        string _denotation;   // Обозначение заказа
        string _creationDate; // Дата формирования
        int _countItems;      // Количество изделий
        string _comment;      // Комментарий обработки
        bool isCalculated;    // Признак обсчета заказа 

        public int ProductId
        {
            get { return _productId; }
            set { _productId = value; }
        }

        public int ObjectId
        {
            get { return _objectId; }
            set { _objectId = value; }
        }

        public string UnitCode
        {
            get { return _unitCode; }
            set { _unitCode = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string Denotation
        {
            get { return _denotation; }
            set { _denotation = value; }
        }

        public string CreationDate
        {
            get { return _creationDate; }
            set { _creationDate = value; }
        }

        public int CountItems
        {
            get { return _countItems; }
            set { _countItems = value; }
        }

        public string Comment
        {
            get { return _comment; }
            set { _comment = value; }
        }

        public bool IsCalculated
        {
            get { return isCalculated; }
            set { isCalculated = value; }
        }
    }

    // Классы T-FLEX (Id)
    public static class TFDClass
    {
        public static readonly int Complex = 502;
        public static readonly int Complement = 501;
        public static readonly int OtherItem = 566;
        public static readonly int StandartItem = 510;
        public static readonly int Assembly = 500;
        public static readonly int Detail = 499;
        public static readonly int Material = 839;

        public static readonly string ComplementName = "Комплекты";
        public static readonly string OtherItemsName = "Прочие изделия";
        public static readonly string StandartItemsName = "Стандартные изделия";
        public static readonly string AssemblyName = "Сборочные единицы";
        public static readonly string DetailName = "Детали";
        public static readonly string MaterialName = "Материалы";

        public static string InsertSection(int classId)
        {
            string sectionName = string.Empty;
            if (classId == Complement)
                sectionName = ComplementName;
            if (classId == OtherItem)
                sectionName = OtherItemsName;
            if (classId == StandartItem)
                sectionName = StandartItemsName;
            if (classId == Assembly)
                sectionName = AssemblyName;
            if (classId == Detail)
                sectionName = DetailName;
            if (classId == Material)
                sectionName = MaterialName;
            return sectionName;
        }

    }
    public delegate void LogDelegate(string line);

    public class ReportParameters
    {



        // типы объектов
        bool documentation; //Документация
        bool complex; //Комплексы
        bool detail; //Детали
        bool assembly; //Сборочные единицы
        bool assemblyCode; // шифрованные 
        bool assemblyCodeless; // нешифрованные 
        bool standart; //Стандартные изделия
        bool other; //Прочие изделия
        bool materials; //Материалы
        bool complement; //Комплекты
        bool complexProgram; //Комплексы (программы)
        bool componentProgram; //Компоненты (программы) 

        public bool AssemblyCodeless
        {
            get { return assemblyCodeless; }
            set { assemblyCodeless = value; }
        }
        public bool AssemblyCode
        {
            get { return assemblyCode; }
            set { assemblyCode = value; }
        }
        public bool Documentation
        {
            get { return documentation; }
            set { documentation = value; }
        }

        public bool Complex
        {
            get { return complex; }
            set { complex = value; }
        }

        public bool Detail
        {
            get { return detail; }
            set { detail = value; }
        }

        public bool Assembly
        {
            get { return assembly; }
            set { assembly = value; }
        }

        public bool Standart
        {
            get { return standart; }
            set { standart = value; }
        }

        public bool Other
        {
            get { return other; }
            set { other = value; }
        }

        public bool Materials
        {
            get { return materials; }
            set { materials = value; }
        }

        public bool Complement
        {
            get { return complement; }
            set { complement = value; }
        }

        public bool ComplexProgram
        {
            get { return complexProgram; }
            set { complexProgram = value; }
        }

        public bool ComponentProgram
        {
            get { return componentProgram; }
            set { componentProgram = value; }
        }

    }

    public class TFDDocument
    {
        // Параметры объекта

        public int ObjectID;
        public int ParentID;
        public int Class;
        public double Amount;
        public string Zone;
        public int Position;
        public string Format;
        public string Naimenovanie;
        public string Denotation;
        public string Remarks;
        public int Level;
        public string Letter;
        public string Code;
        public double TotalCount;
        public int productId;

        public TFDDocument(TFDDocument doc)
        {
            this.ObjectID = doc.ObjectID;
            this.ParentID = doc.ParentID;
            this.Class = doc.Class;
            this.Amount = doc.Amount;
            this.Zone = doc.Zone;
            this.Position = doc.Position;
            this.Format = doc.Format;
            this.Naimenovanie = doc.Naimenovanie;
            this.Denotation = doc.Denotation;
            this.Remarks = doc.Remarks;
            this.Level = doc.Level;
            this.Letter = doc.Letter;
            this.Code = doc.Code;
            this.TotalCount = doc.TotalCount;
        }
        public TFDDocument()
        {
        }

    }


    internal class NaimenAndOboznachComparer : IComparer<TFDDocument>
    {
        public int Compare(TFDDocument ob1, TFDDocument ob2)
        {
            string designation1 = ob1.Denotation.Replace(" ", "");
            string designation2 = ob2.Denotation.Replace(" ", "");
            int ob = String.Compare(designation1, designation2);
            if (ob != 0)
                return ob;
            else
                return String.Compare(ob1.Naimenovanie, ob2.Naimenovanie);
        }
    }
    public static class SQLRequest
    {
        public static readonly string ComplexWithCode = @"SELECT DISTINCT n.s_ObjectID,
                                                 n.Denotation,
                                                 n.Name,
                                                ISNULL(duc.Code,'') as UnitCode,
                                        
                                                 n.s_ClassID
                         
	                                   FROM Nomenclature n
                                             LEFT JOIN UnitCodes duc ON (n.s_ObjectID = duc.s_ObjectID)
	                                   WHERE n.s_Deleted = 0 AND n.s_ActualVersion = 1 AND
                                          duc.s_Deleted  in( 0,null) AND duc.s_ActualVersion in (1, null) AND	                     
	                                         n.s_ClassID in ({0})
	                                   ORDER BY  n.Denotation,
	                                   n.s_ClassID,
	                                            
                                                 n.Name,
                                                  n.s_ObjectID,
                                                  UnitCode
                                                
                                                 ";


        public static readonly string NameComplexWithCode = @"SELECT n.Name, n.Denotation, duc.Code AS UnitCode 
                                                              FROM Nomenclature n LEFT JOIN UnitCodes duc ON n.s_ObjectID = duc.s_ObjectID
                                                              WHERE n.s_Deleted = 0 AND n.s_ActualVersion = 1 AND
                                                                    duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 AND
                                                                    n.s_ObjectID = {0}";

        public static readonly string TotalCountDocument = @"DECLARE @docid INT
                                                                DECLARE @productId INT
                                                                DECLARE @level INT
                                                                DECLARE @insertcount INT
                                                                DECLARE @CollectionID INT
                                                                SET @docid = {0} -- @rootID -- 886421
                                                                SET @insertcount = 0
                                                                SET @level=0

                                                                IF OBJECT_ID('tempdb..#TmpR') IS NULL
                                                                CREATE TABLE #TmpR (s_ObjectID INT,             
                                                                                      s_ParentID INT, 
                                                                                      [level] INT,
                                                                                      s_ClassID INT,                      
                                                                                      Name NVARCHAR(255),
                                                                                      Denotation NVARCHAR(150),     
                                                                                      Amount FLOAT,
                                                                                      UnitCode NVARCHAR(255) NULL,
                                                                                      TotalCount FLOAT,
                                                                                      Remarks NVARCHAR(255)) 
                                                                ELSE TRUNCATE TABLE #TmpR

                                                                INSERT INTO #TmpR
                                                                SELECT DISTINCT n.s_ObjectID, -1, 0, n.s_ClassID,
                                                                        n.Name, n.Denotation,  1, duc.Code ,1,''
                                                                FROM Nomenclature n INNER JOIN Documents d ON (d.s_NomenclatureObjectID = n.s_ObjectID)
                                                                                    LEFT JOIN UnitCodes duc ON (n.s_ObjectID = duc.s_ObjectID)                                       
                                                                WHERE n.s_ObjectID=@docid AND
                                                                      n.s_Deleted=0 AND n.s_ActualVersion=1 AND
                                                                      d.s_Deleted = 0 AND d.s_ActualVersion = 1 AND    
                                                                      duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 
    
    
                                                                WHILE 1=1
                                                                BEGIN

                                                                   INSERT INTO #TmpR
                                                                   SELECT nh.s_ObjectID, nh.s_ParentID, @level+1,
                                                                          n.s_ClassID, n.Name, n.Denotation, 
                                                                          nh.Amount,  ISNULL(duc.Code,'') as UnitCode,#TmpR.TotalCount*nh.Amount, nh.Remarks
                                                                   FROM NomenclatureHierarchy nh INNER JOIN #TmpR ON nh.s_ParentID = #TmpR.s_ObjectID AND
                                                                                                                       #TmpR.[level] = @level AND
                                                                                                                       nh.s_ActualVersion = 1 AND nh.s_Deleted = 0
                                                                                                 INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID AND
                                                                                                                              n.s_ActualVersion = 1 AND n.s_Deleted = 0 AND
                                                                                                                              n.s_ClassID IN ({1})
                                                                                                                              -- Сборочная единица, Комплект, 
                                                                                                                              -- Прочее изделие, Стандартное изделие, Материалы
                                                                                                 LEFT JOIN Documents d ON nh.s_ObjectID = d.s_NomenclatureObjectID AND
                                                                                                                          d.s_Deleted=0 AND d.s_ActualVersion=1
                                                                                                 LEFT JOIN UnitCodes duc ON n.s_ObjectID = duc.s_ObjectID AND
                                                                                                                            duc.s_Deleted = 0 AND duc.s_ActualVersion = 1
                                
                                                                    SET @insertcount = @@ROWCOUNT 
                                                                    SET @level = @level + 1

                                                                    IF  @insertcount = 0
                                                                    GOTO END1
                                                                END             	
                                                                END1:

                                                                --ORDER BY Denotation, Name	
                                                               SELECT #TmpR.s_ObjectID,                       
                                                                                      #TmpR.s_ClassID,                      
                                                                                      #TmpR.Name,
                                                                                      #TmpR.Denotation,        
                                                                                      #TmpR.UnitCode,
                                                                                      Sum (TotalCount*{2}) as summa,
                                                                                      #TmpR.Remarks,
                                                                                      #TmpR.Amount
                                                                FROM #TmpR
                                                                GROUP BY   #TmpR.s_ClassID, #TmpR.s_ObjectID, #TmpR.Denotation, #TmpR.Name,                  
                                                                                      #TmpR.UnitCode, #TmpR.Remarks,#TmpR.Amount";

        public static readonly string ComplexSelected = @"SELECT n.Name, n.Denotation, '' AS UnitCode 
                                                              FROM Nomenclature n 
                                                              WHERE n.s_Deleted = 0 AND n.s_ActualVersion = 1 AND
                                                                    n.s_ObjectID = {0}";
    }

    public class PurchaseDesignProductOperation
    {
        // Соединение с БД T-FLEX DOCs
        public static SqlConnection GetConnectionTFLEX()
        {
            SqlConnectionStringBuilder stringBuilder = new SqlConnectionStringBuilder();
            var info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            var reference = info.CreateReference();
            var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
            stringBuilder.DataSource = nameSqlServer;
            stringBuilder.InitialCatalog = "TFlexDOCs";
            stringBuilder.Password = "reportUser";
            stringBuilder.UserID = "reportUser";
            string connectionString = stringBuilder.ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);

            return connection;
        }

        // Получение шифрованных заказов, доступных для обсчета
        public static List<ItemInfo> GetProducts(string classToString)
        {
            List<ItemInfo> complexes = new List<ItemInfo>();

            using (SqlConnection connection = GetConnectionTFLEX())
            {
                connection.Open();

                string complexQuery = String.Format(SQLRequest.ComplexWithCode, classToString);

                SqlCommand complexCommand = new SqlCommand(complexQuery, connection);
                complexCommand.CommandTimeout = 0;
                SqlDataReader reader = complexCommand.ExecuteReader();

                while (reader.Read())
                {
                    ItemInfo complex = new ItemInfo();
                    complex.Id = Convert.ToInt32(reader[0]);
                    complex.Denotation = Convert.ToString(reader[1]);
                    complex.Name = Convert.ToString(reader[2]);
                    complex.UnitCode = Convert.ToString(reader[3]);
                    complex.ClassId = Convert.ToInt32(reader[4]);

                    complexes.Add(complex);
                }
                reader.Close();
            }

            return complexes;
        }

        // Сложение количеств повторяющихся элементов отсортированной коллекции
        static private List<TFDDocument> SumDublicates(List<TFDDocument> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].ObjectID == data[slaveID].ObjectID)
                    //    if (data[mainID].Naimenovanie == data[slaveID].Naimenovanie)
                    {
                        data[mainID].TotalCount += data[slaveID].TotalCount;
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
            return data;
        }

        // Построение дерева заказа
        public List<TFDDocument> BuildTree(int docId, int countItem, ReportParameters reportParameters)
        {
            List<TFDDocument> totalTree = SumDublicates(GetTotalTree(docId, countItem, reportParameters));

            return totalTree;
        }
        public List<TFDDocument> listComplex = new List<TFDDocument>();


        List<string> listIncludeName = new List<string>();

        private void AddItemIncludeName()
        {
            listIncludeName.Add("Геркон МК");            //1.Геркон
            listIncludeName.Add("Диод 2А536А-5");        //2.Диод
            listIncludeName.Add("Диод 2А536А-6");
            listIncludeName.Add("Диод 2Д419Б");
            listIncludeName.Add("Диод 2Д429А");
            listIncludeName.Add("Конденсатор К10-71-2"); //3.Конденсатор
            listIncludeName.Add("Конденсатор К78-2");
            listIncludeName.Add("Микросхема 132");       //4.Микросхема
            listIncludeName.Add("Микросхема 1108");
            listIncludeName.Add("Микросхема 1485");
            listIncludeName.Add("Микросхема 1564");
            listIncludeName.Add("Микросхема 1821");
            listIncludeName.Add("Микросхема 301");
            listIncludeName.Add("Миксросхема 544 УД");
            listIncludeName.Add("Микросхема 564");
            listIncludeName.Add("Набор резисторов НР1-20"); //5.Набор резисторов
            listIncludeName.Add("Резсистор Р1-6");          //6.Резистор
            listIncludeName.Add("Резистор С5-53");
            listIncludeName.Add("Резистор СП5-2");
            listIncludeName.Add("Резистор СП5-16");
            listIncludeName.Add("Блок Б19К");             //7.Блок
            listIncludeName.Add("Терморезистор СТ3");      //8.Терморезистор
            listIncludeName.Add("Терморезистор ММТ");
            listIncludeName.Add("Резонатор РГ");           //9.Резонатор
            listIncludeName.Add("Резонатор К1");
            listIncludeName.Add("Стабилиторн 2С133");      //10.Стабилитрон
            listIncludeName.Add("Стабилиторн 2С147");
            listIncludeName.Add("Стабилиторн 2С168");
            listIncludeName.Add("Стабилиторн Д818");
            listIncludeName.Add("Транзистор 3П612");       //11.Транзистор
            listIncludeName.Add("Транзистор 2Т643");
            listIncludeName.Add("Транзистор 2Т644");
            listIncludeName.Add("Транзистор 2Т866");
            listIncludeName.Add("Транзистор 2П762А");
            listIncludeName.Add("Транзистор 2Т974А");
            listIncludeName.Add("Устройство согласующее УС-1");    //12.Устройство согласующее УС-1
            listIncludeName.Add("Гарнитура со средней шумозащитой");  //13.Гарнитура со средней шумозащитой


        }

        // Вставить информацию о комплексе в таблицы Products и Objects (инициализация обсчета)
        public List<TFDDocument> GetTotalTree(int docId, int countItem, ReportParameters reportParameters)
        {

            // int lastId;
            List<TFDDocument> totalTree;
            // TFDDocument selectedComplex = new TFDDocument();

            using (SqlConnection connection = GetConnectionTFLEX())
            {
                connection.Open();
                string usingTypes = TFDClass.Assembly.ToString() + "," + TFDClass.Complement.ToString();

                if (reportParameters.Detail)
                {
                    usingTypes += "," + TFDClass.Detail.ToString();
                }

                if (reportParameters.Materials)
                {
                    usingTypes += "," + TFDClass.Material.ToString();
                }

                if (reportParameters.Other)
                {
                    usingTypes += "," + TFDClass.OtherItem.ToString();
                }

                if (reportParameters.Standart)
                {
                    usingTypes += "," + TFDClass.StandartItem.ToString();
                }


                // получение параметров комплекса
                string totalCountRequest = String.Format(SQLRequest.TotalCountDocument, docId, usingTypes, countItem);
                SqlCommand totalCountCommand = new SqlCommand(totalCountRequest, connection);
                totalCountCommand.CommandTimeout = 0;
                SqlDataReader readerTotalCount = totalCountCommand.ExecuteReader();

                totalTree = new List<TFDDocument>();
                AddItemIncludeName();

                while (readerTotalCount.Read())
                {
                    TFDDocument document = new TFDDocument();
                    document.ObjectID = GetSqlInt32(readerTotalCount, 0);
                    document.Class = GetSqlInt32(readerTotalCount, 1);

                    //   document.Naimenovanie = nameReadItem;
                    document.Naimenovanie = GetSqlString(readerTotalCount, 2);
                    document.Denotation = GetSqlString(readerTotalCount, 3);
                    document.Code = GetSqlString(readerTotalCount, 4);
                    document.TotalCount = GetSqlDouble(readerTotalCount, 5);
                    document.productId = docId;


                    if ((TFDClass.OtherItem == document.Class) && (document.Naimenovanie.Contains("Резистор") || document.Naimenovanie.Contains("Конденсатор")))
                    {
                        double amountByRemark = CalculateAmountByRemark(GetSqlString(readerTotalCount, 6));
                        if (GetSqlDouble(readerTotalCount, 7) > amountByRemark)
                            amountByRemark = GetSqlDouble(readerTotalCount, 7);

                        document.TotalCount = amountByRemark * GetSqlDouble(readerTotalCount, 5) / GetSqlDouble(readerTotalCount, 7);
                    }

                    if (GetSqlInt32(readerTotalCount, 0) == docId)
                    {
                        if (document.ObjectID == docId)
                            listComplex.Add(document);
                        continue;
                    }



                    if (GetSqlInt32(readerTotalCount, 1) == TFDClass.Assembly)
                    {
                        if (reportParameters.Assembly)
                        {
                            if (reportParameters.AssemblyCode)
                            {
                                if (GetSqlString(readerTotalCount, 4) != string.Empty)
                                {
                                    totalTree.Add(document);
                                    continue;

                                }
                                if (reportParameters.AssemblyCodeless)
                                {
                                    totalTree.Add(document);
                                }
                            }
                            else
                            {
                                if (reportParameters.AssemblyCodeless)
                                {
                                    if (GetSqlString(readerTotalCount, 4) == string.Empty)
                                    {
                                        totalTree.Add(document);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (GetSqlInt32(readerTotalCount, 1) == TFDClass.Complement)
                        {

                            if (reportParameters.Complement)
                                totalTree.Add(document);

                        }
                        else
                        {
                            totalTree.Add(document);
                        }
                    }
                }
                readerTotalCount.Close();
            }
            return totalTree;
        }

        public int CalculateAmountByRemark(string remark)
        {
            int amount = 0;
            string threeDotsPattern = @"[RCС]\d{1,3}\.{2,3}[RCС]\d{1,3}";
            string digitsPattern = @"\d{1,3}";
            string itemsPattern = @"[RCС]\d{1,3}";

            // Вычисление количества в структуре типа RX..RZ amount = (Z-X)+1
            var threeDotsFragments = Regex.Matches(remark, threeDotsPattern);
            foreach (Match match in threeDotsFragments)
            {
                var digitsFragments = Regex.Matches(match.Value, digitsPattern);
                if (Convert.ToInt32(digitsFragments[0].Value) < Convert.ToInt32(digitsFragments[1].Value))
                    amount += Convert.ToInt32(digitsFragments[1].Value) - Convert.ToInt32(digitsFragments[0].Value) + 1;
                else
                    amount += Convert.ToInt32(digitsFragments[0].Value) - Convert.ToInt32(digitsFragments[1].Value) + 1;
                remark = remark.Remove(remark.IndexOf(match.Value), match.Value.Length);
            }

            // Вычисление количества одиночных элементов
            amount += Regex.Matches(remark, itemsPattern).Count;
            return amount;
        }

        private static string GetSqlString(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return String.Empty;

            return reader.GetString(field);
        }

        private static int GetSqlInt32(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0;

            return reader.GetInt32(field);
        }

        private static double GetSqlDouble(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0d;

            return reader.GetDouble(field);
        }
    }
}
