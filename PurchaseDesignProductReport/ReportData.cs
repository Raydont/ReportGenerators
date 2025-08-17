using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;

namespace PurchaseDesignProductReport
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
        public static readonly int Complekt = 501;
        public static readonly int OtherItem = 566;
        public static readonly int StandartItem = 510;
        public static readonly int Assembly = 500;
        public static readonly int Detail = 499;
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
            stringBuilder.IntegratedSecurity = true;
            stringBuilder.PersistSecurityInfo = false;
            string connectionString = stringBuilder.ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);

            return connection;
        }

        // Соединение с БД отчета
        public static SqlConnection GetConnectionPKI()
        {
            SqlConnectionStringBuilder stringBuilder = new SqlConnectionStringBuilder();
            var info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            var reference = info.CreateReference();
            var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
            stringBuilder.DataSource = nameSqlServer;
            stringBuilder.InitialCatalog = "PKIReportDB";
            stringBuilder.UserID = "gsuser";
            stringBuilder.Password = "V64m548";
            stringBuilder.IntegratedSecurity = false;
            stringBuilder.PersistSecurityInfo = false;
            string connectionString = stringBuilder.ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);

            return connection;
        }

        // Получение заказов, доступных для обсчета
        public static List<ItemInfo> GetProducts()
        {
            List<ItemInfo> complexes = new List<ItemInfo>();

            using (SqlConnection connection = GetConnectionTFLEX())
            {
                connection.Open();

                string complexQuery = String.Format(
                 @"SELECT DISTINCT n.s_ObjectID,
                             n.Denotation,
                             n.Name,
                             ISNULL(duc.Code,'') as UnitCode,
                             n.s_ClassID
	               FROM Nomenclature n
                        LEFT JOIN UnitCodes duc ON (n.s_ObjectID = duc.s_ObjectID)
	               WHERE n.s_Deleted = 0 AND n.s_ActualVersion = 1 AND
                         duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 AND	                     
	                     n.s_ClassID = {0} -- ID класса Комплекс
	               ORDER BY n.Denotation, n.Name", TFDClass.Complex);
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

        // Построение дерева заказа
        public static void BuildTree(int docId, bool isPKIReport)
        {
            int productId = InsertProduct(docId);
            FillProduct(docId, productId, isPKIReport);
        }

        // Вставить информацию о комплексе в таблицы Products и Objects (инициализация обсчета)
        public static int InsertProduct(int docId)
        {
            string name = string.Empty;
            string denotation = string.Empty;
            string unitCode = string.Empty;
            int lastId;

            using (SqlConnection connection = GetConnectionTFLEX())
            {
                connection.Open();

                // получение параметров комплекса
                string complexParamQuery = String.Format(
                    @"SELECT n.Name, n.Denotation, duc.Code AS UnitCode 
                      FROM Nomenclature n LEFT JOIN UnitCodes duc ON n.s_ObjectID = duc.s_ObjectID
                      WHERE n.s_Deleted = 0 AND n.s_ActualVersion = 1 AND
                            duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 AND
                            n.s_ObjectID = {0}", docId);
                SqlCommand complexParamCommand = new SqlCommand(complexParamQuery, connection);
                complexParamCommand.CommandTimeout = 0;
                SqlDataReader reader = complexParamCommand.ExecuteReader();

                while (reader.Read())
                {
                    name = reader.GetString(0);
                    denotation = reader.GetString(1);
                    unitCode = reader.GetString(2);
                }
                reader.Close();
            }

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string insertProductQuery = String.Format(
                    @"DECLARE @lastId INT
                      DECLARE @denotation NVARCHAR(150)
                      DECLARE @name NVARCHAR(255)
                      DECLARE @unitCode NVARCHAR(255)
                      DECLARE @docId INT

                      SET @denotation = N'{0}'
                      SET @name = N'{1}'
                      SET @unitCode = N'{2}'
                      SET @docId = {3}

                      INSERT Products(DocID, Name, Denotation, UnitCode, CreationDate, Comment, Calculated)
                      VALUES (@docId, @name, @denotation, @unitCode, GETDATE(), @unitCode + ' Объект вставлен', 0);

                      SET @lastId = @@IDENTITY;
                  
                      INSERT INTO Objects(ProductID, Designation, Name, Material, Path, UnitCode, Type, TotalCount, DocID, ParentID)
                      VALUES(@lastId, @denotation, @name, NULL, NULL, @unitCode, 0, 1, @docId, -1);                      

                      SELECT @lastId;", denotation, name, unitCode, docId);
                SqlCommand insertProductCommand = new SqlCommand(insertProductQuery, connection);
                insertProductCommand.CommandTimeout = 0;
                lastId = Convert.ToInt32(insertProductCommand.ExecuteScalar());
            }

            return lastId;
        }

        // Формирование структуры заказа
        public static void FillProduct(int objectId, int productId, bool isPKIReport)
        {
            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_STRUCTURE_FORMING, "Запуск процесса формирования структуры заказа");
            FillProductStructure(objectId, productId, isPKIReport);
            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_STRUCTURE_FORMED, "Структура заказа сформирована");
        }

        // Добавление записи в лог для заказа
        public static void AddProductsLogRecord(int productId, int status, string comment)
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string addProductsLogRecordQuery = String.Format(
                    @"INSERT INTO ProductsLog (ProductID, Status, RecordDate, Comment)
                  VALUES ({0}, {1}, GetDate(), N'{2}')", productId, status, comment);
                SqlCommand addProductsLogRecordCommand = new SqlCommand(addProductsLogRecordQuery, connection);
                addProductsLogRecordCommand.CommandTimeout = 0;
                addProductsLogRecordCommand.ExecuteNonQuery();
            }
        }

        // Добавление дерева заказа в БД обсчета
        public static void FillProductStructure(int docId, int productId, bool isPKIReport)
        {
            List<ItemInfo> complexItems = new List<ItemInfo>();
            int lastObjectId;
            bool occurenceInserted;
            int itemType = 13;

            #region Вставка объектов в составе заказа в Objects
            complexItems = GetTreeByItems(docId, isPKIReport);
            foreach (ItemInfo complexItem in complexItems)
            {
                if (complexItem.ClassId == TFDClass.OtherItem) itemType = (int)ItemType.TYPE_PKI;       // Прочие изделия
                if (complexItem.ClassId == TFDClass.StandartItem) itemType = (int)ItemType.TYPE_PKI;    // Стандартные изделия
                if (complexItem.ClassId == TFDClass.Assembly) itemType = (int)ItemType.TYPE_ASSEMBLY;   // Сборочная единица
                if (complexItem.ClassId == TFDClass.Detail) itemType = (int)ItemType.TYPE_DETAIL;       // Деталь

                lastObjectId = InsertObject(productId, complexItem.Denotation, complexItem.Name, complexItem.Material, complexItem.Marshrut, complexItem.UnitCode, itemType, complexItem.Id, complexItem.ParentId);
            }
            #endregion

            #region Вставка вхождений объектов в составе заказа в Occurence
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                foreach (ItemInfo complexItem in complexItems)
                {
                    if (complexItem.Id == -1) continue;

                    occurenceInserted = InsertOccurence(productId, complexItem.ParentId, complexItem.Id, (int)complexItem.Amount);
                }
            }
            #endregion
        }

        // Получение элементов дерева заказа из структуры T-FLEX DOCs
        public static List<ItemInfo> GetTreeByItems(int rootId, bool isPKIReport)
        {
            List<ItemInfo> items = new List<ItemInfo>();

            string classIds;

            if (isPKIReport)
                classIds = TFDClass.Assembly + ", " + TFDClass.Complekt + ", " + TFDClass.OtherItem + ", " + TFDClass.StandartItem;
            else
                classIds = TFDClass.Assembly + ", " + TFDClass.Complekt + ", " + TFDClass.Detail;

            using (SqlConnection connection = GetConnectionTFLEX())
            {
                connection.Open();

                #region Получение элементов дерева заказа
                string elementsTreeQuery = String.Format(
                      @"DECLARE @docid INT
                        DECLARE @productId INT
                        DECLARE @level INT
                        DECLARE @insertcount INT
                        DECLARE @CollectionID INT
                        SET @docid = {0} -- @rootID -- 886421
                        SET @insertcount = 0
                        SET @level=0

                        IF OBJECT_ID('tempdb..#TmpRep') IS NULL
                        CREATE TABLE #TmpRep (s_ObjectID INT,             
                                              s_ParentID INT, 
                                              [level] INT,
                                              s_ClassID INT,                      
                                              Name NVARCHAR(255),
                                              Denotation NVARCHAR(150),
                                              Amount FLOAT,
                                              Marshrut NVARCHAR(255) NULL,
                                              Material NVARCHAR(255) NULL,
                                              BomSection NVARCHAR(255),
                                              UnitCode NVARCHAR(255) NULL) 
                        ELSE TRUNCATE TABLE #TmpRep
                  
                        INSERT INTO #TmpRep
                        SELECT DISTINCT n.s_ObjectID, -1, 0, n.s_ClassID,
                                n.Name, n.Denotation, 1, 
                                tp.Marshrut, tp.Material,
                                N'Комплексы', duc.Code
                        FROM Nomenclature n INNER JOIN Documents d ON (d.s_NomenclatureObjectID = n.s_ObjectID)
                                            LEFT JOIN Table_432 tp ON d.s_ObjectID = tp.s_ObjectID,
                                            UnitCodes duc                                       
                        WHERE n.s_ObjectID=@docid AND
                            n.s_Deleted=0 AND n.s_ActualVersion=1 AND
                            n.s_ObjectID = duc.s_ObjectID AND
                            duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 --AND
                            --d.s_ObjectID = tp.s_ObjectID

                        WHILE 1=1
                        BEGIN
   
                           INSERT INTO #TmpRep
                           SELECT DISTINCT nh.s_ObjectID, nh.s_ParentID, @level+1,
                                  n.s_ClassID, n.Name, n.Denotation,
                                  nh.Amount, tp.Marshrut, tp.Material,
                                  nh.BOMSection, ISNULL(duc.Code,'') as UnitCode
                           FROM NomenclatureHierarchy nh INNER JOIN #TmpRep ON nh.s_ParentID = #TmpRep.s_ObjectID AND
                                                                               #TmpRep.[level] = @level AND
                                                                               nh.s_ActualVersion = 1 AND nh.s_Deleted = 0
                                                         INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID AND
                                                                                      n.s_ActualVersion = 1 AND n.s_Deleted = 0 AND
                                                                                      n.s_ClassID IN ({1})
                                                                                      -- Сборочная единица, Комплект, 
                                                                                      -- Прочее изделие, Стандартное изделие
                                                         LEFT JOIN Documents d ON nh.s_ObjectID = d.s_NomenclatureObjectID AND
                                                                                  d.s_Deleted=0 AND d.s_ActualVersion=1
                                                         LEFT JOIN UnitCodes duc ON n.s_ObjectID = duc.s_ObjectID AND
                                                                                    duc.s_Deleted = 0 AND duc.s_ActualVersion = 1
                                                         LEFT JOIN Table_432 tp ON d.s_ObjectID = tp.s_ObjectID AND
                                                                                   tp.s_Deleted = 0 AND tp.s_ActualVersion = 1

                            SET @insertcount = @@ROWCOUNT 
                            SET @level = @level + 1
    
                            IF  @insertcount = 0
                            GOTO END1
                        END             	
                        END1:

                        SELECT * FROM #TmpRep
                        ORDER BY level", rootId, classIds);
                SqlCommand elementsTreeCommand = new SqlCommand(elementsTreeQuery, connection);
                elementsTreeCommand.CommandTimeout = 0;
                SqlDataReader reader = elementsTreeCommand.ExecuteReader();
                #endregion

                while (reader.Read())
                {
                    ItemInfo item = new ItemInfo();
                    item.Id = reader.GetInt32(0);
                    item.ParentId = reader.GetInt32(1);
                    item.Level = reader.GetInt32(2);
                    item.ClassId = reader.GetInt32(3);
                    item.Name = reader.GetString(4);
                    item.Denotation = reader.GetString(5);
                    item.Amount = reader.GetDouble(6);
                    item.Marshrut = (reader[7] != DBNull.Value ? reader.GetString(7) : "");
                    item.Material = (reader[8] != DBNull.Value ? reader.GetString(8) : "");
                    item.BomSection = reader.GetString(9);
                    item.UnitCode = reader.GetString(10);

                    items.Add(item);
                }
                reader.Close();
            }

            return items;
        }

        // Вставка объекта с параметрами в таблицу Objects
        public static int InsertObject(int productId, string code, string name, string material, string path, string unitCode, int type, int objectId, int parentId)
        {
            int id;

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                // получение объекта из таблицы Objects
                string getObjectQuery = String.Format(
                    @"SELECT ID FROM Objects 
                      WHERE DocID = {0} AND ProductID = {1}", objectId, productId);
                SqlCommand getObjectQueryCommand = new SqlCommand(getObjectQuery, connection);
                getObjectQueryCommand.CommandTimeout = 0;
                if (getObjectQueryCommand.ExecuteScalar() != null)
                {
                    id = Convert.ToInt32(getObjectQueryCommand.ExecuteScalar());
                    return id;
                }

                string insertObjectQuery = String.Format(
                    @"INSERT INTO Objects(ProductID, Designation, Name, Material, Path, UnitCode, Type, DocID, ParentID)
                      VALUES({0}, '{1}', '{2}', '{3}', '{4}', '{5}', {6}, {7}, {8})

                      SELECT @@IDENTITY",
                      productId, code, name, material, path, unitCode, type, objectId, parentId);
                SqlCommand insertObjectCommand = new SqlCommand(insertObjectQuery, connection);
                insertObjectCommand.CommandTimeout = 0;
                id = Convert.ToInt32(insertObjectCommand.ExecuteScalar());
            }

            return id;
        }

        // Вставка вхождения объекта в состав другого в таблице вхождений Occurence
        public static bool InsertOccurence(int productId, int parentId, int objectId, int counter)
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string findOccurenceQuery = String.Format(
                    @"SELECT ID FROM Occurence 
                      WHERE ProductID={0} AND ObjectID={1} AND ParentObjectID={2}", productId, objectId, parentId);
                SqlCommand findOccurenceCommand = new SqlCommand(findOccurenceQuery, connection);
                findOccurenceCommand.CommandTimeout = 0;
                if (findOccurenceCommand.ExecuteScalar() != null) return false;

                string insertOccurenceQuery = String.Format(
                    @"INSERT INTO Occurence (ProductID, ObjectID, ParentObjectID, Count) VALUES ({0}, {1}, {2}, {3})",
                      productId, objectId, parentId, counter);
                SqlCommand insertOccurenceCommand = new SqlCommand(insertOccurenceQuery, connection);
                insertOccurenceCommand.CommandTimeout = 0;
                insertOccurenceCommand.ExecuteNonQuery();
            }

            return true;
        }

        // Получение информации о добавленных заказах
        public static List<ComplexInfo> GetAddedProducts()
        {
            List<ComplexInfo> complexInfos = new List<ComplexInfo>();

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string getComplexInfoQuery = @"SELECT ID, DocID, Name, Denotation, UnitCode, CreationDate, Comment, Calculated 
                                               FROM Products
                                               ORDER BY CreationDate desc";
                SqlCommand getComplexInfoCommand = new SqlCommand(getComplexInfoQuery, connection);
                getComplexInfoCommand.CommandTimeout = 0;
                SqlDataReader reader = getComplexInfoCommand.ExecuteReader();

                while (reader.Read())
                {
                    ComplexInfo complexInfo = new ComplexInfo();
                    complexInfo.ProductId = reader.GetInt32(0);
                    complexInfo.ObjectId = reader.GetInt32(1);
                    complexInfo.Name = reader.GetString(2);
                    complexInfo.Denotation = reader.GetString(3);
                    complexInfo.UnitCode = reader.GetString(4);
                    complexInfo.CreationDate = reader.GetDateTime(5).ToShortDateString();
                    complexInfo.CountItems = GetProductObjectCount(complexInfo.ProductId);
                    complexInfo.Comment = reader.GetString(6);
                    complexInfo.IsCalculated = Convert.ToInt32(reader[7]) == 0 ? false : true;

                    complexInfos.Add(complexInfo);
                }
                reader.Close();

                return complexInfos;
            }
        }

        // Получение количества объектов в заказе
        public static int GetProductObjectCount(int productId)
        {
            int count;

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string getCountQuery = String.Format(
                    @"SELECT COUNT(*) FROM Objects WHERE ProductID = {0}", productId);
                SqlCommand getCountCommand = new SqlCommand(getCountQuery, connection);
                getCountCommand.CommandTimeout = 0;
                count = (int)getCountCommand.ExecuteScalar();
            }

            return count;
        }

        // Обсчет заказа
        public static void Calculate(int prodObjectId)
        {
            int totalCount = 1;        // начальное значение для общего количества любого из элементов
            int productId;

            // получение ProductID добавленного комплекса
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string getProductIdQuery = String.Format(
                    @"SELECT ID FROM Products WHERE Calculated=0 AND DocID={0}", prodObjectId);
                SqlCommand getProductIdCommand = new SqlCommand(getProductIdQuery, connection);
                getProductIdCommand.CommandTimeout = 0;
                productId = (int)getProductIdCommand.ExecuteScalar();
            }

            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_CALCULATING, "Расчет суммарного количества изделий");
            CalculateTotalCount(productId, prodObjectId, totalCount, prodObjectId, 1, 1);
            SetProductCalculated(productId);
            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_CALCULATED, "Расчет количества изделий заказа завершен");
        }

        // Увеличение общего количества объектов на величину totalCount
        public static void UpdateTotalCount(int productId, int objectId, int totalCount)
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string updateTotalCountQuery = String.Format(
                    @"UPDATE Objects SET TotalCount = TotalCount + {0} WHERE ProductId = {1} AND DocID = {2}", totalCount, productId, objectId);
                SqlCommand updateTotalCountCommand = new SqlCommand(updateTotalCountQuery, connection);
                updateTotalCountCommand.CommandTimeout = 0;
                updateTotalCountCommand.ExecuteNonQuery();
            }
        }

        // Вычисление общего количества (рекурсивно)
        public static void CalculateTotalCount(int productId, int objectId, int totalCount, int unitId, int calcUnitCount, int unitCountMultiply)
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                CalculateTotalCount(productId, objectId, totalCount, unitId, calcUnitCount, unitCountMultiply, connection);
            }
        }

        public static void CalculateTotalCount(int productId, int objectId, int totalCount, int unitId, int calcUnitCount, int unitCountMultiply, SqlConnection connection)
        {
            double newtotcount,
                   nextCalcUnitCount,
                   nextUnitCountMultiply;
            int curUnitId;
            List<ItemInfo> itemsInfo = new List<ItemInfo>();

            // считываем потомков объекта objectId
            string query = String.Format(
                @"SELECT oc.ObjectID, oc.Count, ob.UnitCode, ob.UnitCalculated
                      FROM Occurence oc INNER JOIN Objects ob ON (oc.ObjectID = ob.DocID AND oc.ProductID = ob.ProductID)
                      WHERE oc.ParentObjectID = {0} AND ob.ProductID = {1}", objectId, productId);
            SqlCommand command = new SqlCommand(query, connection);
            command.CommandTimeout = 0;
            SqlDataReader reader = command.ExecuteReader();

            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    ItemInfo item = new ItemInfo();
                    item.Id = Convert.ToInt32(reader[0]);
                    item.Amount = Convert.ToInt32(reader[1]);
                    item.UnitCode = reader[2].ToString();
                    item.UnitCalculated = Convert.ToInt32(reader[3]);

                    itemsInfo.Add(item);
                }
            }
            reader.Close();

            // если потомки есть, то выполняем для них обсчет
            if (itemsInfo.Count > 0)
            {
                for (int i = 0; i < itemsInfo.Count; i++)
                {
                    newtotcount = totalCount * itemsInfo[i].Amount; //MessageBox.Show("Новое количество для id =" + itemsInfo[i].Id + " равно " + newtotcount.ToString());
                    UpdateTotalCount(productId, itemsInfo[i].Id, (int)newtotcount); //MessageBox.Show("Обновление для id = " + itemsInfo[i].Id + " завершено");
                    // закончили обсчет количества в пределах заказа

                    // если шифр пуст
                    if (itemsInfo[i].UnitCode.Trim() == "")
                    {
                        // не шифрованная
                        curUnitId = unitId;
                        nextCalcUnitCount = calcUnitCount;
                        nextUnitCountMultiply = unitCountMultiply * itemsInfo[i].Amount;
                    }
                    else
                    {
                        // шифрованная
                        curUnitId = itemsInfo[i].Id;
                        nextCalcUnitCount = itemsInfo[i].UnitCalculated == 0 ? 1 : 0;
                        nextUnitCountMultiply = 1;
                    }

                    // установка признака обсчета для элемента c шифром
                    if (itemsInfo[i].UnitCode.Trim() != "")
                    {
                        string updateQuery = String.Format(
                            @"UPDATE Objects SET UnitCalculated=1 
                              WHERE DocID={0} AND ProductID ={1}", itemsInfo[i].Id, productId);
                        SqlCommand updateCommand = new SqlCommand(updateQuery, connection);
                        updateCommand.CommandTimeout = 0;
                        updateCommand.ExecuteNonQuery();
                    }

                    CalculateTotalCount(productId, itemsInfo[i].Id, (int)newtotcount, curUnitId, (int)nextCalcUnitCount, (int)nextUnitCountMultiply, connection);
                } //for
            } //if
        }

        // Установка признака обсчета заказа
        public static void SetProductCalculated(int productId)
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string updatePartCalculatedQuery = String.Format(
                    @"UPDATE Products SET Calculated=1 WHERE ID = {0}", productId);
                SqlCommand updatePartCalculatedCommand = new SqlCommand(updatePartCalculatedQuery, connection);
                updatePartCalculatedCommand.CommandTimeout = 0;
                updatePartCalculatedCommand.ExecuteNonQuery();
            }
        }

        // Получение номенклатуры объектов для отчета и их общего количества для всех комплексов отчета
        public static List<ItemInfo> GetReportData(List<int> productIds, int type)
        {
            List<ItemInfo> items = new List<ItemInfo>();

            //MessageBox.Show(productIds.Count.ToString());

            string products = string.Empty;
            for (int i = 0; i < productIds.Count; i++)
            {
                products += productIds[i];
                if (i != productIds.Count - 1) products += ", ";
            }

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                // получение номенклатуры объектов
                string getNomenclatureQuery = String.Format(
                    @"SELECT DocID, Designation, Name, SUM(TotalCount) as TotalCount
                      FROM [Objects]
                      WHERE (ProductID IN ({0}) and Type={1})
                      GROUP BY DocID, Designation, Name
                      ORDER BY Designation, Name", products, type);
                SqlCommand getNomenclatureCommand = new SqlCommand(getNomenclatureQuery, connection);
                getNomenclatureCommand.CommandTimeout = 0;
                SqlDataReader reader = getNomenclatureCommand.ExecuteReader();

                while (reader.Read())
                {
                    ItemInfo item = new ItemInfo();
                    item.Id = reader.GetInt32(0);
                    item.Denotation = reader[1] != DBNull.Value ? reader.GetString(1) : string.Empty;
                    item.Name = reader.GetString(2);
                    item.Amount = reader.GetInt32(3);

                    items.Add(item);
                }
            }

            return items;
        }

        // Получение параметров вхождения объекта для комплекса
        public static List<ItemInfo> GetObjectOccurenceParams(int objectId, int productId)
        {
            List<ItemInfo> itemParams = new List<ItemInfo>();

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string getObjectOccurenceParamsQuery = String.Format(
                    @"SELECT oc.ParentObjectID, oc.Count
                      FROM Occurence oc INNER JOIN [Objects] ob ON (oc.ObjectID = ob.DocID AND oc.ProductID = ob.ProductID)
                      WHERE oc.ObjectID = {0} AND ob.ProductID = {1} 
                      ORDER BY Count", objectId, productId);
                SqlCommand getObjectOccurenceParamsCommand = new SqlCommand(getObjectOccurenceParamsQuery, connection);
                getObjectOccurenceParamsCommand.CommandTimeout = 0;
                SqlDataReader reader = getObjectOccurenceParamsCommand.ExecuteReader();

                while (reader.Read())
                {
                    ItemInfo itemInfo = new ItemInfo();
                    itemInfo.ParentId = reader.GetInt32(0);
                    itemInfo.Amount = reader.GetInt32(1);

                    itemParams.Add(itemInfo);
                }
                reader.Close();
            }

            return itemParams;
        }

        // Получение параметров объекта
        public static ItemInfo GetObjectParams(int objectId, int productId)
        {
            ItemInfo itemInfo;

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string getObjectParamsQuery = String.Format(
                    @"SELECT Name, Designation, UnitCode, TotalCount FROM Objects WHERE DocID={0} AND ProductID={1}", objectId, productId);
                SqlCommand getObjectParamsCommand = new SqlCommand(getObjectParamsQuery, connection);
                getObjectParamsCommand.CommandTimeout = 0;
                SqlDataReader reader = getObjectParamsCommand.ExecuteReader();

                itemInfo = new ItemInfo();
                while (reader.Read())
                {
                    itemInfo.Name = reader.GetString(0);
                    itemInfo.Denotation = reader.GetString(1);
                    itemInfo.UnitCode = reader.GetString(2);
                    itemInfo.Amount = reader.GetInt32(3);
                }
            }

            return itemInfo;
        }

        // Получение общего количества объекта в заказе
        public static int GetObjectTotalCountInProduct(int objectId, int productId)
        {
            int count;

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string getObjectTotalCountQuery = String.Format(
                    @"SELECT TotalCount FROM Objects
                      WHERE DocID={0} AND ProductID={1}", objectId, productId);
                SqlCommand getObjectTotalCountCommand = new SqlCommand(getObjectTotalCountQuery, connection);
                getObjectTotalCountCommand.CommandTimeout = 0;
                count = Convert.ToInt32(getObjectTotalCountCommand.ExecuteScalar());
            }

            return count;
        }

        // Получение информации о комплексе
        public static ComplexInfo GetProductInfo(int productId)
        {
            ComplexInfo complexInfo;

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string productInfoQuery = String.Format(
                    @"SELECT DocID, Name, Denotation, UnitCode, CreationDate, Comment, Calculated
                      FROM Products
                      WHERE ID={0}", productId);
                SqlCommand productInfoCommand = new SqlCommand(productInfoQuery, connection);
                productInfoCommand.CommandTimeout = 0;
                SqlDataReader reader = productInfoCommand.ExecuteReader();

                complexInfo = new ComplexInfo();
                while (reader.Read())
                {
                    complexInfo.ObjectId = reader.GetInt32(0);
                    complexInfo.Name = reader.GetString(1);
                    complexInfo.Denotation = reader.GetString(2);
                    complexInfo.UnitCode = reader.GetString(3);
                    complexInfo.CreationDate = reader.GetDateTime(4).ToShortDateString();
                    complexInfo.Comment = reader.GetString(5);
                    complexInfo.IsCalculated = (Convert.ToInt32(reader[6]) == 0 ? false : true);
                }
                reader.Close();

                return complexInfo;
            }
        }

        // Очистка таблиц добавленных заказов
        public static void DeleteAllAddedProducts()
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string clearTablesQuery =
                      @"USE PKIReportDB;
                        BEGIN TRANSACTION;
                        TRUNCATE TABLE dbo.Products;
                        TRUNCATE TABLE dbo.ProductsLog;
                        TRUNCATE TABLE dbo.Objects;
                        TRUNCATE TABLE dbo.Occurence;                        
                        COMMIT TRANSACTION;";
                SqlCommand clearTablesCommand = new SqlCommand(clearTablesQuery, connection);
                clearTablesCommand.CommandTimeout = 0;
                clearTablesCommand.ExecuteNonQuery();
            }
        }

        // Удаление структуры заказа
        public static void DeleteProductStructure(int productId)
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string deleteProductStructureQuery = String.Format(
                    @"BEGIN TRANSACTION;
                      DELETE FROM Occurence WHERE ProductID = {0};
                      DELETE FROM Objects WHERE ProductID = {0};                      
                      DELETE FROM Products WHERE ID = {0};
                      DELETE FROM ProductsLog WHERE ProductID = {0};
                      COMMIT TRANSACTION;", productId);
                SqlCommand deleteProductStructureCommand = new SqlCommand(deleteProductStructureQuery, connection);
                deleteProductStructureCommand.CommandTimeout = 0;
                deleteProductStructureCommand.ExecuteNonQuery();
            }
        }
    }
}
