using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PurchaseDesignProductReportNoRecursive
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
        double _totalCount;            // Поле для расчета общего количества элемента
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

        public double TotalCount
        {
            get { return _totalCount; }
            set { _totalCount = value; }
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
        int _productId;       // ID в таблице ProductsWR (добавленных заказов)
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
            stringBuilder.DataSource = TFlex.DOCs.Model.ServerGateway.Connection.ServerName;
            stringBuilder.InitialCatalog = "TFlexDOCs";
            stringBuilder.Password = "reportUser";
            stringBuilder.UserID = "reportUser";
            stringBuilder.MinPoolSize = 10;
            string connectionString = stringBuilder.ConnectionString;
            SqlConnection connection = new SqlConnection(connectionString);

            return connection;
        }

        // Соединение с БД отчета
        public static SqlConnection GetConnectionPKI()
        {
            SqlConnectionStringBuilder stringBuilder = new SqlConnectionStringBuilder();
            stringBuilder.DataSource = TFlex.DOCs.Model.ServerGateway.Connection.ServerName;
            stringBuilder.InitialCatalog = "PKIReportDB";
            stringBuilder.UserID = "gsuser";
            stringBuilder.Password = "V64m548";
            stringBuilder.IntegratedSecurity = false;
            stringBuilder.PersistSecurityInfo = false;
            stringBuilder.MinPoolSize = 10;
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
        public static bool BuildTree(int docId, bool isPKIReport, bool isFirstUnitCodeDevices, string userFullName, int count)
        {
            int productId = InsertProduct(docId);
            if (CheckWorkingUser(userFullName))
                return true;
            else
            {
                FillProduct(docId, productId, isPKIReport, count);
                if (isFirstUnitCodeDevices) FillFirstUnitCodeDevicesTable(docId, productId, count);
                return false;
            }
        }

        public static void DeleteUser()
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();
                string deleteUserQuery = String.Format(@"DELETE FROM [PKIReportDB].[dbo].[UserNR]");
                var deleteUserCommand = new SqlCommand(deleteUserQuery, connection);
                deleteUserCommand.CommandTimeout = 0;
                deleteUserCommand.ExecuteNonQuery();
            }

        }

        // Вставить информацию о комплексе в таблицы ProductsWR и ObjectsWR (инициализация обсчета)
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

                      INSERT ProductsNR(DocID, Name, Denotation, UnitCode, CreationDate, Comment, Calculated)
                      VALUES (@docId, @name, @denotation, @unitCode, GETDATE(), @unitCode + ' Объект вставлен', 0);

                      SET @lastId = @@IDENTITY;
                  
                      INSERT INTO ObjectsNR(ProductID, Designation, Name, Material, Path, UnitCode, Type, TotalCount, DocID, ParentID)
                      VALUES(@lastId, @denotation, @name, NULL, NULL, @unitCode, 0, 1, @docId, -1);                      

                      SELECT @lastId;", denotation, name, unitCode, docId);
                SqlCommand insertProductCommand = new SqlCommand(insertProductQuery, connection);
                insertProductCommand.CommandTimeout = 0;
                lastId = Convert.ToInt32(insertProductCommand.ExecuteScalar());
            }

            return lastId;
        }

        // Формирование структуры заказа
        public static void FillProduct(int objectId, int productId, bool isPKIReport, int count)
        {
            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_STRUCTURE_FORMING, "Запуск процесса формирования структуры заказа и расчета количества изделий");
            FillProductStructure(objectId, productId, isPKIReport, count);
            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_STRUCTURE_FORMED, "Расчет количества изделий заказа завершен");

        }

        // Добавление записи в лог для заказа
        public static void AddProductsLogRecord(int productId, int status, string comment)
        {


            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string addProductsLogRecordQuery = String.Format(
                    @"INSERT INTO ProductsLogNR (ProductID, Status, RecordDate, Comment)
                      VALUES ({0}, {1}, GetDate(), N'{2}')", productId, status, comment);
                SqlCommand addProductsLogRecordCommand = new SqlCommand(addProductsLogRecordQuery, connection);
                addProductsLogRecordCommand.CommandTimeout = 0;
                addProductsLogRecordCommand.ExecuteNonQuery();
            }
        }

        public static bool CheckWorkingUser(string userFullName)
        {
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string getUserLoginQuery = String.Format(@"SELECT users.[Login]
                                                      FROM [PKIReportDB].[dbo].[UserNR] workUserId INNER JOIN [TFLEXDOCs].[dbo].[UserParameters] users 
                                                      ON (workUserId.UserId = users.s_ObjectID)");
                var objectLogintCommand = new SqlCommand(getUserLoginQuery, connection);
                objectLogintCommand.CommandTimeout = 0;
                var loginUser = objectLogintCommand.ExecuteScalar();

                if (loginUser != null)
                {
                    if (loginUser.ToString().Trim() != userFullName.Trim())
                    {
                        string getUserNameQuery = String.Format(@"SELECT users.[FullName]
                                                      FROM [PKIReportDB].[dbo].[UserNR] workUserId INNER JOIN [TFLEXDOCs].[dbo].[Users] users 
                                                      ON (workUserId.UserId = users.s_ObjectID)");
                        var objectLNameCommand = new SqlCommand(getUserNameQuery, connection);
                        objectLNameCommand.CommandTimeout = 0;
                        var fullNameUser = objectLNameCommand.ExecuteScalar();
                        MessageBox.Show(
                            "В данный момент запустить отчет нельзя, поскольку с базой данных работает " + fullNameUser,
                            "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    string addProductsLogRecordQuery = String.Format(
                   @"DECLARE @UserId NvarChar(255)
                     SET @UserId = (SELECT u.s_ObjectID
                     FROM [TFLEXDOCs].[dbo].[Users] u INNER JOIN [TFLEXDOCs].[dbo].[UserParameters] up ON (u.s_ObjectID = up.s_ObjectID)
                     WHERE up.[Login] =  N'{0}' AND u.s_ActualVersion = 1 AND u.s_Deleted = 0)
                     INSERT INTO [PKIReportDB].[dbo].[UserNR] (UserId)
                     VALUES (@UserId)", userFullName);
                    var addProductsLogRecordCommand = new SqlCommand(addProductsLogRecordQuery, connection);
                    addProductsLogRecordCommand.CommandTimeout = 0;
                    addProductsLogRecordCommand.ExecuteNonQuery();
                    return false;
                }
            }
        }


        // Добавление дерева заказа в БД обсчета
        public static void FillProductStructure(int docId, int productId, bool isPKIReport, int count)
        {
            List<ItemInfo> complexItems;
            int lastObjectId;
            int itemType = 13;
            //bool occurenceInserted;

            #region Вставка объектов в составе заказа в Objects
            complexItems = GetTreeByItems(docId, isPKIReport, count);
            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();
                foreach (ItemInfo complexItem in complexItems)
                {
                    if (complexItem.ClassId == TFDClass.OtherItem) itemType = (int)ItemType.TYPE_PKI;       // Прочие изделия
                    if (complexItem.ClassId == TFDClass.StandartItem) itemType = (int)ItemType.TYPE_PKI;    // Стандартные изделия
                    if (complexItem.ClassId == TFDClass.Assembly) itemType = (int)ItemType.TYPE_ASSEMBLY;   // Сборочная единица
                    if (complexItem.ClassId == TFDClass.Detail) itemType = (int)ItemType.TYPE_DETAIL;       // Деталь

                    lastObjectId = InsertObject(connection, productId, complexItem.Denotation, complexItem.Name, complexItem.Material, complexItem.Marshrut, complexItem.UnitCode, itemType, (int)complexItem.Amount, (int)complexItem.TotalCount, complexItem.Id, complexItem.ParentId);
                }
            }
            #endregion

            #region Вставка вхождений объектов в составе заказа в Occurence (было для рекурсивной версии)
            //using (SqlConnection connection = GetConnectionPKI())
            //{
            //    connection.Open();

            //    foreach (ItemInfo complexItem in complexItems)
            //    {
            //        if (complexItem.Id == -1) continue;

            //        occurenceInserted = InsertOccurence(productId, complexItem.ParentId, complexItem.Id, (int)complexItem.Amount);
            //    }
            //}
            #endregion
        }

        // Получение элементов дерева заказа из структуры T-FLEX DOCs
        public static List<ItemInfo> GetTreeByItems(int rootId, bool isPKIReport, int count)
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
                                              UnitCode NVARCHAR(255) NULL,
                                              TotalCount FLOAT,
                                              Remarks NVARCHAR(255)) 
                        ELSE TRUNCATE TABLE #TmpRep
                  
                        INSERT INTO #TmpRep
                        SELECT DISTINCT n.s_ObjectID, -1, 0, n.s_ClassID,
                                n.Name, n.Denotation, 1, 
                                tp.Marshrut, tp.Material,
                                N'Комплексы', duc.Code, 1,''
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
                           SELECT nh.s_ObjectID, nh.s_ParentID, @level+1,
                                  n.s_ClassID, n.Name, n.Denotation,
                                  nh.Amount, tp.Marshrut, tp.Material,
                                  nh.BOMSection, ISNULL(duc.Code,'') as UnitCode, #TmpRep.TotalCount*nh.Amount, nh.Remarks
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
                    item.TotalCount = count * reader.GetDouble(11);

                    if ((TFDClass.OtherItem == item.ClassId) && (item.Name.Contains("Резистор") || item.Name.Contains("Конденсатор")))
                    {
                        double amountByRemark = CalculateAmountByRemark(reader.GetString(12));

                        if (reader.GetDouble(6) > amountByRemark)
                            amountByRemark = reader.GetDouble(6);
                        else
                            item.Amount = amountByRemark;

                        item.TotalCount = amountByRemark * reader.GetDouble(11) / reader.GetDouble(6);
                    }

                    items.Add(item);
                }
                reader.Close();
            }

            return items;
        }

        public static int CalculateAmountByRemark(string remark)
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

        static SqlCommand insertObjectCommand;
        static SqlConnection insertObjectLastConnection;

        // Вставка объекта с параметрами в таблицу Objects
        public static int InsertObject(SqlConnection connection, int productId, string code, string name, string material, string path, string unitCode, int type, int count, int totalCount, int objectId, int parentId)
        {
            int id;



            #region Было для рекурсивной версии
            //                // получение объекта из таблицы Objects
            //                string getObjectQuery = String.Format(
            //                    @"SELECT ID FROM Objects 
            //                      WHERE DocID = {0} AND ProductID = {1}", objectId, productId);
            //                SqlCommand getObjectQueryCommand = new SqlCommand(getObjectQuery, connection);
            //                getObjectQueryCommand.CommandTimeout = 0;
            //                if (getObjectQueryCommand.ExecuteScalar() != null)
            //                {
            //                    id = Convert.ToInt32(getObjectQueryCommand.ExecuteScalar());
            //                    return id;
            //                }
            #endregion

            if (insertObjectCommand == null || connection != insertObjectLastConnection)
            {
                insertObjectLastConnection = connection;

                const string insertObjectQuery =
                    @"INSERT INTO ObjectsNR(ProductID, Designation, Name, Material, Path, UnitCode, Type, Count, TotalCount, DocID, ParentID)
                      VALUES(@productId, @code, @name, @material, @path, @unitCode, @type, @count, @totalCount, @objectId, @parentId)

                      SELECT @@IDENTITY";


                insertObjectCommand = new SqlCommand(insertObjectQuery, connection);
                insertObjectCommand.Parameters.Add(new SqlParameter("@productId", System.Data.SqlDbType.Int));
                insertObjectCommand.Parameters.Add(new SqlParameter("@code", System.Data.SqlDbType.NVarChar, 255));
                insertObjectCommand.Parameters.Add(new SqlParameter("@name", System.Data.SqlDbType.NVarChar, 1024));
                insertObjectCommand.Parameters.Add(new SqlParameter("@material", System.Data.SqlDbType.NVarChar, 1024));
                insertObjectCommand.Parameters.Add(new SqlParameter("@path", System.Data.SqlDbType.NVarChar, 4000));
                insertObjectCommand.Parameters.Add(new SqlParameter("@unitCode", System.Data.SqlDbType.NVarChar, 255));
                insertObjectCommand.Parameters.Add(new SqlParameter("@type", System.Data.SqlDbType.Int));
                insertObjectCommand.Parameters.Add(new SqlParameter("@count", System.Data.SqlDbType.Int));
                insertObjectCommand.Parameters.Add(new SqlParameter("@totalCount", System.Data.SqlDbType.Int));
                insertObjectCommand.Parameters.Add(new SqlParameter("@objectId", System.Data.SqlDbType.Int));
                insertObjectCommand.Parameters.Add(new SqlParameter("@parentId", System.Data.SqlDbType.Int));

                insertObjectCommand.CommandTimeout = 0;
                insertObjectCommand.Prepare();
            }

            insertObjectCommand.Parameters[0].Value = productId;
            insertObjectCommand.Parameters[1].Value = code;
            insertObjectCommand.Parameters[2].Value = name;
            insertObjectCommand.Parameters[3].Value = material;
            insertObjectCommand.Parameters[4].Value = path;
            insertObjectCommand.Parameters[5].Value = unitCode;
            insertObjectCommand.Parameters[6].Value = type;
            insertObjectCommand.Parameters[7].Value = count;
            insertObjectCommand.Parameters[8].Value = totalCount;
            insertObjectCommand.Parameters[9].Value = objectId;
            insertObjectCommand.Parameters[10].Value = parentId;


            id = Convert.ToInt32(insertObjectCommand.ExecuteScalar());


            return id;
        }

        // Вставка объекта с параметрами в таблицу FirstUnitCodeDevicesNR
        public static int InsertFirstUnitCodeDevice(int productId, int objectId, string unitCode, string denotation, int count)
        {
            int id;

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();

                string insertFirstUnitCodeDeviceQuery = String.Format(
                    @"INSERT INTO FirstUnitCodeDevicesNR(ProductID, DocID, Code, Designation, Count)
                      VALUES({0}, {1}, '{2}', '{3}', {4})

                      SELECT @@IDENTITY",
                      productId, objectId, unitCode, denotation, count);
                SqlCommand insertFirstUnitCodeDeviceCommand = new SqlCommand(insertFirstUnitCodeDeviceQuery, connection);
                insertFirstUnitCodeDeviceCommand.CommandTimeout = 0;
                id = Convert.ToInt32(insertFirstUnitCodeDeviceCommand.ExecuteScalar());
            }

            return id;
        }

        // Получение списка первых шифрованных устройств, куда входит список деталей, для заказа
        public static List<ItemInfo> GetFirstUnitCodeDevicesList(int rootId, int productId, int count)
        {
            List<ItemInfo> items = new List<ItemInfo>();

            using (SqlConnection connection = GetConnectionTFLEX())
            {
                connection.Open();

                string classIds = TFDClass.Detail + ", " + TFDClass.Assembly + ", " + TFDClass.Complekt;
                #region Поиск списка первых шифрованных устройств, в которые входят детали
                string firstUnitCodeDevicesQuery = String.Format(
                  @"DECLARE @docid INT                   -- s_ObjectID заказа             
                    DECLARE @level INT                   -- уровень вложенности элемента
                    DECLARE @insertcount INT             -- количество вставленных строк        
                    DECLARE @Denotation NVARCHAR(255)    -- обозначение первого шифрованного устройства
                    DECLARE @Code NVARCHAR(255)          -- шифр первого шифрованного устройства
                    DECLARE @ObjectID INT                -- s_ObjectID объекта детали

                    SET @docid =  {0}                    -- rootId                   
                    SET @insertcount = 0                 
                    SET @level=0                            

                    -- 1. Заполнение таблицы дерева структуры заказа

                    -- таблица дерева структуры заказа
                    IF OBJECT_ID('tempdb..#TmpRep') IS NULL
                    CREATE TABLE #TmpRep (ID INT IDENTITY(1,1),
                                          s_ObjectID INT,             
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

                    -- инициализация таблицы дерева структуры заказа корневым элементом заказа
                    INSERT INTO #TmpRep
                    SELECT DISTINCT n.s_ObjectID, -1, 0, n.s_ClassID,
                            n.Name, n.Denotation, 1, 
                            tp.Marshrut, tp.Material,
                            N'Комплексы', duc.Code--, 1
                    FROM Nomenclature n INNER JOIN Documents d ON (d.s_NomenclatureObjectID = n.s_ObjectID)
                                        LEFT JOIN Table_432 tp ON d.s_ObjectID = tp.s_ObjectID,
                                        UnitCodes duc                                       
                    WHERE n.s_ObjectID=@docid AND
                        n.s_Deleted=0 AND n.s_ActualVersion=1 AND
                        n.s_ObjectID = duc.s_ObjectID AND
                        duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 --AND
                        --d.s_ObjectID = tp.s_ObjectID

                    -- заполнение таблицы дерева структуры заказа
                    WHILE 1=1
                    BEGIN

                       INSERT INTO #TmpRep
                       SELECT DISTINCT nh.s_ObjectID, nh.s_ParentID, @level+1,
                              n.s_ClassID, n.Name, n.Denotation,
                              nh.Amount, tp.Marshrut, tp.Material,
                              nh.BOMSection, ISNULL(duc.Code,'') as UnitCode--, #TmpRep.TotalCount*nh.Amount
                       FROM NomenclatureHierarchy nh INNER JOIN #TmpRep ON nh.s_ParentID = #TmpRep.s_ObjectID AND
                                                                           #TmpRep.[level] = @level AND
                                                                           nh.s_ActualVersion = 1 AND nh.s_Deleted = 0
                                                     INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID AND
                                                                                  n.s_ActualVersion = 1 AND n.s_Deleted = 0 AND
                                                                                  n.s_ClassID IN ({1})
                                                                                  -- Деталь, Сборочная единица, Комплект                                                              
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

                    -- 2. Заполнение таблицы всех деталей, входящих в заказ

                    -- таблица всех деталей, входящих в заказ
                    IF OBJECT_ID('tempdb..#Details') is NULL
                    CREATE TABLE #Details (ID INT IDENTITY(1,1),
                                           s_ObjectID INT)
                    ELSE TRUNCATE TABLE #Details

                    -- заполнение таблицы деталей, входящих в заказ
                    INSERT INTO #Details
                    SELECT DISTINCT s_ObjectID
                    FROM #TmpRep 
                    WHERE s_ClassID = {2}

                    -- 3. Поиск первого шифрованного устройства в заказе

                    DECLARE @CountDetails INT      -- Количество элементов в таблице деталей заказа
                    DECLARE @CounterDetails INT    -- Счетчик цикла по таблице деталей заказа


                    -- таблица деталей с одинаковым s_ObjectID для цикла поиска шифрованных устройств
                    IF OBJECT_ID('tempdb..#Item') is NULL
                                    CREATE TABLE #Item (ID int identity(1,1),s_ObjectID int,s_ParentID int, Amount FLOAT)
                                    ELSE
                                    TRUNCATE TABLE #Item

                    -- таблица нешифрованных устройств для цикла поиска шифрованных устройств
                    IF OBJECT_ID('tempdb..#NoUCodeItem') is NULL
                                    CREATE TABLE #NoUCodeItem (s_ObjectID int,s_ParentID int, Amount FLOAT)
                                    ELSE
                                    TRUNCATE TABLE #NoUCodeItem
                
                    -- Итоговая таблица деталей с первым шифрованным устройством
                    IF OBJECT_ID('tempdb..#Total') is NULL
                                    CREATE TABLE #Total (s_ObjectID_Detail int, s_ParentID_FUDevice int, Code nvarchar(255), Denotation nvarchar(255), Amount FLOAT)
                                    ELSE
                                    TRUNCATE TABLE #Total


                    DECLARE @Count INT            -- Количество элементов в таблице #Item (переменное)
                    DECLARE @Counter INT          -- Счетчик цикла по элементам таблицы #Item
                    DECLARE @ObjID INT            -- ID детали в составе комплекса
                    DECLARE @PObjID INT           -- ID родителя детали объекта в составе комплекса
                    DECLARE @ICount INT           -- Количество элементов в таблице #Item
                    DECLARE @PObjectID INT        -- ID родителя элемента в таблице #Item
                    DECLARE @Amount FLOAT         -- Количество объекта в таблице #Item

                    -- количество деталей в таблице #Details
                    SET @CountDetails = (SELECT COUNT(*) FROM #Details)
                    SET @CounterDetails = 1

                    WHILE @CounterDetails <= @CountDetails
                    BEGIN
   
	                    -- ID объекта в составе комплекса
	                    SET @ObjID = (SELECT s_ObjectID FROM #Details WHERE ID=@CounterDetails)
   
	                    -- очищение таблицы #Item и запись в нее данных, соответствующих выбранному устройству в составе комплекса
	                    TRUNCATE TABLE #Item
	                    INSERT INTO #Item
	                    SELECT s_ObjectID, s_ParentID, Amount FROM #TmpRep
	                    WHERE s_ObjectID = @ObjID	
	
	                    -- цикл поиска первого шифрованного устройства, в которое входит эта деталь
	                    WHILE 1=1
	                    BEGIN
	      
	                          -- количество деталей с s_ObjectID = @ObjID
		                      SET @ICount = (SELECT COUNT(*) FROM #Item)
		                      SET @Counter = 1
	      
		                      -- очистка таблицы нешифрованных устройств
		                      TRUNCATE TABLE #NoUCodeItem
		  
		                      -- проверка наличия шифра
		                      WHILE @Counter <= @ICount
		                      BEGIN
			                     -- ID объекта и объекта-родителя
			                     SET @PObjectID = (SELECT s_ParentID FROM #Item WHERE ID = @Counter)
			                     SET @ObjectID = (SELECT s_ObjectID FROM #Item WHERE ID = @Counter)			 
			 
			                     -- из таблицы UnitCodes вытаскиваем шифр и обозначение объекта-родителя
			                     SET @Code = (SELECT Code FROM UnitCodes
						                      WHERE s_ActualVersion = 1 AND s_Deleted = 0 AND
								                    s_ObjectID = @PObjectID)
			                     SET @Denotation = (SELECT DISTINCT Denotation FROM #TmpRep
			                                        WHERE s_ObjectID = @PObjectID)
			                     SET @Amount = (SELECT TOP(1) Amount FROM #Item
			                                    WHERE s_ObjectID = @ObjectID AND s_ParentID = @PObjectID) 
			 
			                     -- ЕСЛИ шифр для объекта-родителя отсутствует или является пустой строкой, ТО объект-родитель записывается в таблицу #NoUCodeItem,
			                     -- ИНАЧЕ объект-родитель определяется как первое шифрованное устройство для детали
			                     IF @Code IS NULL OR @Code = ''
				                     INSERT INTO #NoUCodeItem
				                     SELECT s_ObjectID, s_ParentID, Amount*@Amount
				                     FROM #TmpRep WHERE s_ObjectID = @PObjectID
			                     ELSE
			                         INSERT INTO #Total
			                         SELECT @ObjID, s_ObjectID, @Code, @Denotation, @Amount
			                         FROM #TmpRep WHERE s_ObjectID = @PObjectID				 
			 
			                     -- увеличение счетчика цикла по одинаковым деталям
			                     SET @Counter = @Counter + 1
		                      END
	      
		                      -- очистка таблицы деталей с одинаковым s_ObjectID
		                      -- и запись в нее данных таблицы #NoUCodeItem
		                      TRUNCATE TABLE #Item 
		                      INSERT INTO #Item
		                      SELECT s_ObjectID, s_ParentID, Amount FROM #NoUCodeItem
	      
		                      -- Проверка, если таблица #Item пустая, то конец цикла
		                      SET @Count = (SELECT COUNT(*) FROM #Item)
		                      IF @Count = 0 GOTO END2

	                    END
	                    END2:
	
	                    -- увеличение счетчика цикла по деталям в заказе
	                    SET @CounterDetails = @CounterDetails + 1
   
                    END
                    
                    -- 4. Получение таблицы первых шифрованных устройств                    
                    SELECT DISTINCT s_ObjectID_Detail, Code, Denotation, Amount
                    FROM #Total
                    ORDER BY s_ObjectID_Detail", rootId, classIds, TFDClass.Detail);
                SqlCommand firstUnitCodeDevicesCommand = new SqlCommand(firstUnitCodeDevicesQuery, connection);
                firstUnitCodeDevicesCommand.CommandTimeout = 0;
                SqlDataReader reader = firstUnitCodeDevicesCommand.ExecuteReader();

                while (reader.Read())
                {
                    ItemInfo itemInfo = new ItemInfo();
                    itemInfo.Id = reader.GetInt32(0);
                    itemInfo.UnitCode = reader.GetString(1);
                    itemInfo.Denotation = reader.GetString(2);
                    itemInfo.Amount = count * reader.GetDouble(3);

                    items.Add(itemInfo);
                }
                #endregion
            }

            return items;
        }

        // Заполнение таблицы первых шифрованных устройств для деталей FirstUnitCodeDevicesNR
        public static void FillFirstUnitCodeDevicesTable(int rootId, int productId, int count)
        {
            List<ItemInfo> firstUnitCodeDevices;
            int lastObjectId;

            #region Вставка элемента первого шифрованного устройства для детали в FirstUnitCodeDevicesNR для данного заказа
            firstUnitCodeDevices = GetFirstUnitCodeDevicesList(rootId, productId, count);
            foreach (ItemInfo firstUnitCodeDevice in firstUnitCodeDevices)
            {
                lastObjectId = InsertFirstUnitCodeDevice(productId, firstUnitCodeDevice.Id, firstUnitCodeDevice.UnitCode, firstUnitCodeDevice.Denotation, (int)firstUnitCodeDevice.Amount);
            }
            #endregion
        }




        // Получение списка первых шифрованных устройств, куда входит данная деталь, для заказа
        public static List<ItemInfo> GetFirstUnitCodeDevicesForDetail(SqlConnection connection, int objectId, int productId)
        {
            List<ItemInfo> items = new List<ItemInfo>();

            //using (SqlConnection connection = GetConnectionPKI())
            {
                //connection.Open();

                string getFirstUnitCodeDevicesQuery = String.Format(
                    @"SELECT DISTINCT Code, Designation, Count 
FROM FirstUnitCodeDevicesNR
WHERE DocID = {0} AND ProductID = {1}
ORDER BY Count", objectId, productId);
                SqlCommand getFirstUnitCodeDevicesCommand = new SqlCommand(getFirstUnitCodeDevicesQuery, connection);
                getFirstUnitCodeDevicesCommand.CommandTimeout = 0;
                SqlDataReader reader = getFirstUnitCodeDevicesCommand.ExecuteReader();

                while (reader.Read())
                {
                    ItemInfo itemInfo = new ItemInfo();
                    itemInfo.UnitCode = reader.GetString(0);
                    itemInfo.Denotation = reader.GetString(1);
                    itemInfo.Amount = reader.GetInt32(2);

                    items.Add(itemInfo);
                }
            }

            return items;
        }

        // Получение информации о добавленных заказах        
        public static List<ComplexInfo> GetAddedProducts()
        {
            List<ComplexInfo> complexInfos = new List<ComplexInfo>();

            using (SqlConnection connection = GetConnectionPKI())
            {
                connection.Open();
                using (SqlConnection connection2 = GetConnectionPKI())
                {
                    connection2.Open();

                    string getComplexInfoQuery = @"SELECT ID, DocID, Name, Denotation, UnitCode, CreationDate, Comment, Calculated 
                                               FROM ProductsNR
                                               ORDER BY CreationDate desc";
                    SqlCommand getComplexInfoCommand = new SqlCommand(getComplexInfoQuery, connection);
                    getComplexInfoCommand.CommandTimeout = 0;


                    string getCountQuery = @"SELECT COUNT(*) FROM ObjectsNR WHERE ProductID = @productID";
                    SqlCommand getCountCommand = new SqlCommand(getCountQuery, connection2);
                    getCountCommand.Parameters.Add(new SqlParameter("@productID", System.Data.SqlDbType.Int));
                    getCountCommand.CommandTimeout = 0;
                    getCountCommand.Prepare();

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

                        getCountCommand.Parameters[0].Value = complexInfo.ProductId;
                        complexInfo.CountItems = (int)getCountCommand.ExecuteScalar();

                        complexInfo.Comment = reader.GetString(6);
                        complexInfo.IsCalculated = Convert.ToInt32(reader[7]) == 0 ? false : true;

                        complexInfos.Add(complexInfo);
                    }
                    reader.Close();

                    return complexInfos;
                }
            }
        }


        #region Рекурсивный расчет общего количества элементов в заказе
        //        // Вставка вхождения объекта в состав другого в таблице вхождений Occurence
        //        public static bool InsertOccurence(int productId, int parentId, int objectId, int counter)
        //        {
        //            using (SqlConnection connection = GetConnectionPKI())
        //            {
        //                connection.Open();

        //                string findOccurenceQuery = String.Format(
        //                    @"SELECT ID FROM Occurence 
        //                              WHERE ProductID={0} AND ObjectID={1} AND ParentObjectID={2}", productId, objectId, parentId);
        //                SqlCommand findOccurenceCommand = new SqlCommand(findOccurenceQuery, connection);
        //                findOccurenceCommand.CommandTimeout = 0;
        //                if (findOccurenceCommand.ExecuteScalar() != null) return false;

        //                string insertOccurenceQuery = String.Format(
        //                    @"INSERT INTO Occurence (ProductID, ObjectID, ParentObjectID, Count) VALUES ({0}, {1}, {2}, {3})",
        //                      productId, objectId, parentId, counter);
        //                SqlCommand insertOccurenceCommand = new SqlCommand(insertOccurenceQuery, connection);
        //                insertOccurenceCommand.CommandTimeout = 0;
        //                insertOccurenceCommand.ExecuteNonQuery();
        //            }

        //            return true;
        //        }


        //        // Обсчет заказа
        //        public static void Calculate(int prodObjectId)
        //        {
        //            int totalCount = 1;        // начальное значение для общего количества любого из элементов
        //            int productId;

        //            // получение ProductID добавленного комплекса
        //            using (SqlConnection connection = GetConnectionPKI())
        //            {
        //                connection.Open();

        //                string getProductIdQuery = String.Format(
        //                    @"SELECT ID FROM Products WHERE Calculated=0 AND DocID={0}", prodObjectId);
        //                SqlCommand getProductIdCommand = new SqlCommand(getProductIdQuery, connection);
        //                getProductIdCommand.CommandTimeout = 0;
        //                productId = (int)getProductIdCommand.ExecuteScalar();
        //            }

        //            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_CALCULATING, "Расчет суммарного количества изделий");
        //            CalculateTotalCount(productId, prodObjectId, totalCount, prodObjectId, 1, 1);
        //            SetProductCalculated(productId);
        //            AddProductsLogRecord(productId, (int)ProductStatus.PRODUCT_STATUS_CALCULATED, "Расчет количества изделий заказа завершен");
        //        }

        //        // Вычисление общего количества (рекурсивно)
        //        public static void CalculateTotalCount(int productId, int objectId, int totalCount, int unitId, int calcUnitCount, int unitCountMultiply)
        //        {
        //            using (SqlConnection connection = GetConnectionPKI())
        //            {
        //                connection.Open();

        //                CalculateTotalCount(productId, objectId, totalCount, unitId, calcUnitCount, unitCountMultiply, connection);
        //            }
        //        }

        //        public static void CalculateTotalCount(int productId, int objectId, int totalCount, int unitId, int calcUnitCount, int unitCountMultiply, SqlConnection connection)
        //        {
        //            double newtotcount,
        //                   nextCalcUnitCount,
        //                   nextUnitCountMultiply;
        //            int curUnitId;
        //            List<ItemInfo> itemsInfo = new List<ItemInfo>();

        //            // считываем потомков объекта objectId
        //            string query = String.Format(
        //                @"SELECT oc.ObjectID, oc.Count, ob.UnitCode, ob.UnitCalculated
        //                      FROM Occurence oc INNER JOIN Objects ob ON (oc.ObjectID = ob.DocID AND oc.ProductID = ob.ProductID)
        //                      WHERE oc.ParentObjectID = {0} AND ob.ProductID = {1}", objectId, productId);
        //            SqlCommand command = new SqlCommand(query, connection);
        //            command.CommandTimeout = 0;
        //            SqlDataReader reader = command.ExecuteReader();

        //            if (reader.HasRows)
        //            {
        //                while (reader.Read())
        //                {
        //                    ItemInfo item = new ItemInfo();
        //                    item.Id = Convert.ToInt32(reader[0]);
        //                    item.Amount = Convert.ToInt32(reader[1]);
        //                    item.UnitCode = reader[2].ToString();
        //                    item.UnitCalculated = Convert.ToInt32(reader[3]);

        //                    itemsInfo.Add(item);
        //                }                
        //            }
        //            reader.Close();

        //            // если потомки есть, то выполняем для них обсчет
        //            if (itemsInfo.Count > 0)
        //            {
        //                for (int i = 0; i < itemsInfo.Count; i++)
        //                {
        //                    newtotcount = totalCount * itemsInfo[i].Amount; //MessageBox.Show("Новое количество для id =" + itemsInfo[i].Id + " равно " + newtotcount.ToString());
        //                    UpdateTotalCount(productId, itemsInfo[i].Id, (int)newtotcount); //MessageBox.Show("Обновление для id = " + itemsInfo[i].Id + " завершено");
        //                    // закончили обсчет количества в пределах заказа

        //                    // если шифр пуст
        //                    if (itemsInfo[i].UnitCode.Trim() == "")
        //                    {                        
        //                        // не шифрованная
        //                        curUnitId = unitId;
        //                        nextCalcUnitCount = calcUnitCount;
        //                        nextUnitCountMultiply = unitCountMultiply * itemsInfo[i].Amount;                        
        //                    }
        //                    else
        //                    {
        //                        // шифрованная
        //                        curUnitId = itemsInfo[i].Id;
        //                        nextCalcUnitCount = itemsInfo[i].UnitCalculated == 0 ? 1 : 0;
        //                        nextUnitCountMultiply = 1;                       
        //                    }

        //                    // установка признака обсчета для элемента c шифром
        //                    if (itemsInfo[i].UnitCode.Trim() != "")
        //                    {
        //                        string updateQuery = String.Format(
        //                            @"UPDATE Objects SET UnitCalculated=1 
        //                              WHERE DocID={0} AND ProductID ={1}", itemsInfo[i].Id, productId);
        //                        SqlCommand updateCommand = new SqlCommand(updateQuery, connection);
        //                        updateCommand.CommandTimeout = 0;
        //                        updateCommand.ExecuteNonQuery();
        //                    }

        //                    CalculateTotalCount(productId, itemsInfo[i].Id, (int)newtotcount, curUnitId, (int)nextCalcUnitCount, (int)nextUnitCountMultiply, connection);
        //                } //for
        //            } //if
        //        }

        //        // Увеличение общего количества объектов на величину totalCount
        //        public static void UpdateTotalCount(int productId, int objectId, int totalCount)
        //        {
        //            using (SqlConnection connection = GetConnectionPKI())
        //            {
        //                connection.Open();

        //                string updateTotalCountQuery = String.Format(
        //                    @"UPDATE Objects SET TotalCount = TotalCount + {0} WHERE ProductId = {1} AND DocID = {2}", totalCount, productId, objectId);
        //                SqlCommand updateTotalCountCommand = new SqlCommand(updateTotalCountQuery, connection);
        //                updateTotalCountCommand.CommandTimeout = 0;
        //                updateTotalCountCommand.ExecuteNonQuery();
        //            }
        //        }

        //// Установка признака обсчета заказа
        //public static void SetProductCalculated(int productId)
        //{
        //    using (SqlConnection connection = GetConnectionPKI())
        //    {
        //        connection.Open();

        //        string updatePartCalculatedQuery = String.Format(
        //            @"UPDATE Products SET Calculated=1 WHERE ID = {0}", productId);
        //        SqlCommand updatePartCalculatedCommand = new SqlCommand(updatePartCalculatedQuery, connection);
        //        updatePartCalculatedCommand.CommandTimeout = 0;
        //        updatePartCalculatedCommand.ExecuteNonQuery();
        //    }
        //}
        #endregion

        // Получение номенклатуры объектов для отчета и их общего количества для всех комплексов отчета
        public static List<ItemInfo> GetReportData(List<int> productIds, int type)
        {
            List<ItemInfo> items = new List<ItemInfo>();

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
                      FROM [ObjectsNR]
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
                    item.TotalCount = reader.GetInt32(3);

                    items.Add(item);
                }
            }

            return items;
        }

        static SqlConnection lastConnectionGetObjectOccurenceParams;
        static SqlCommand getObjectOccurenceParamsCommand;

        // Получение параметров вхождения объекта для комплекса
        public static List<ItemInfo> GetObjectOccurenceParams(SqlConnection connection, int objectId, int productId)
        {
            List<ItemInfo> itemParams = new List<ItemInfo>();

            //using (SqlConnection connection = GetConnectionPKI())
            {
                //connection.Open();

                const string getObjectOccurenceParamsQuery =
@"SELECT DISTINCT ParentID, [Count]
FROM [ObjectsNR]
WHERE DocID = @docID AND ProductID = @productID 
ORDER BY [Count]";

                //, objectId, productId)
                if (getObjectOccurenceParamsCommand == null || connection != lastConnectionGetObjectOccurenceParams)
                {
                    lastConnectionGetObjectOccurenceParams = connection;
                    getObjectOccurenceParamsCommand = new SqlCommand(getObjectOccurenceParamsQuery, connection);
                    getObjectOccurenceParamsCommand.CommandTimeout = 0;
                    getObjectOccurenceParamsCommand.Parameters.Add(new SqlParameter("@docID",  System.Data.SqlDbType.Int) );
                    getObjectOccurenceParamsCommand.Parameters.Add(new SqlParameter("@productID", System.Data.SqlDbType.Int));
                    getObjectOccurenceParamsCommand.Prepare();
                }

                getObjectOccurenceParamsCommand.Parameters[0].Value = objectId;
                getObjectOccurenceParamsCommand.Parameters[1].Value = productId;

                using (SqlDataReader reader = getObjectOccurenceParamsCommand.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ItemInfo itemInfo = new ItemInfo();
                        itemInfo.ParentId = reader.GetInt32(0);
                        itemInfo.Amount = reader.GetInt32(1);

                        itemParams.Add(itemInfo);
                    }
                }
            }

            return itemParams;
        }

        // Получение параметров объекта
        //public static ItemInfo GetObjectParams(int objectId, int productId)
        //{
        //    ItemInfo itemInfo;

        //    using (SqlConnection connection = GetConnectionPKI())
        //    {
        //        connection.Open();

        //        string getObjectParamsQuery = String.Format(
        //            @"SELECT Name, Designation, UnitCode, Count FROM ObjectsNR WHERE DocID={0} AND ProductID={1}", objectId, productId);
        //        SqlCommand getObjectParamsCommand = new SqlCommand(getObjectParamsQuery, connection);
        //        getObjectParamsCommand.CommandTimeout = 0;
        //        SqlDataReader reader = getObjectParamsCommand.ExecuteReader();

        //        itemInfo = new ItemInfo();
        //        while (reader.Read())
        //        {
        //            itemInfo.Name = reader.GetString(0);
        //            itemInfo.Denotation = reader.GetString(1);
        //            itemInfo.UnitCode = reader.GetString(2);
        //            itemInfo.Amount = reader.GetInt32(3);
        //        }
        //    }

        //    return itemInfo;
        //}




        static SqlConnection lastConnectionGetOccurenceParams;
        static SqlCommand getObjectParamsCommand;

        // Получение параметров объектов-вхождений
        public static ItemInfo GetOccurenceParams(SqlConnection connection, int objectId, int productId)
        {
            ItemInfo itemInfo;

            //using (SqlConnection connection = GetConnectionPKI())
            {
                //connection.Open();

                const string getObjectParamsQuery =
@"SELECT Name, Designation, UnitCode, SUM(TotalCount) 
FROM ObjectsNR WHERE DocID=@docID AND ProductID=@productID
GROUP BY Name, Designation, UnitCode";

                if (getObjectParamsCommand == null || connection != lastConnectionGetOccurenceParams)
                {
                    lastConnectionGetOccurenceParams = connection;
                    getObjectParamsCommand = new SqlCommand(getObjectParamsQuery, connection);
                    getObjectParamsCommand.CommandTimeout = 0;
                    getObjectParamsCommand.Parameters.Add(new SqlParameter("@docID", System.Data.SqlDbType.Int));
                    getObjectParamsCommand.Parameters.Add(new SqlParameter("@productID", System.Data.SqlDbType.Int));
                    getObjectParamsCommand.Prepare();
                }

                getObjectParamsCommand.Parameters[0].Value = objectId;
                getObjectParamsCommand.Parameters[1].Value = productId;

                using (SqlDataReader reader = getObjectParamsCommand.ExecuteReader())
                {
                    itemInfo = new ItemInfo();
                    while (reader.Read())
                    {
                        itemInfo.Name = reader.GetString(0);
                        itemInfo.Denotation = reader.GetString(1);
                        itemInfo.UnitCode = reader.GetString(2);
                        itemInfo.TotalCount = reader.GetInt32(3);
                    }
                }
            }

            return itemInfo;
        }

        static SqlConnection lastConnectionGetObjectTotalCountInProduct;
        static SqlCommand getObjectTotalCountCommand;

        // Получение общего количества объекта в заказе
        public static int GetObjectTotalCountInProduct(SqlConnection connection, int objectId, int productId)
        {
            int count;

            const string getObjectTotalCountQuery = @"SELECT SUM(TotalCount) FROM ObjectsNR WHERE DocID=@docID AND ProductID=@productID";

            if (getObjectTotalCountCommand == null || connection != lastConnectionGetObjectTotalCountInProduct)
            {
                lastConnectionGetObjectTotalCountInProduct = connection;
                getObjectTotalCountCommand = new SqlCommand(getObjectTotalCountQuery, connection);
                getObjectTotalCountCommand.CommandTimeout = 0;
                getObjectTotalCountCommand.Parameters.Add(new SqlParameter("@docID", System.Data.SqlDbType.Int));
                getObjectTotalCountCommand.Parameters.Add(new SqlParameter("@productID", System.Data.SqlDbType.Int));
                getObjectTotalCountCommand.Prepare();
            }

            //using (SqlConnection connection = GetConnectionPKI())
            {
                //connection.Open();

                getObjectTotalCountCommand.Parameters[0].Value = objectId;
                getObjectTotalCountCommand.Parameters[1].Value = productId;


                var value = getObjectTotalCountCommand.ExecuteScalar();
                count = (value == DBNull.Value ? 0 : Convert.ToInt32(value));
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
                      FROM ProductsNR
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
                        TRUNCATE TABLE dbo.ProductsNR;
                        TRUNCATE TABLE dbo.ProductsLogNR;
                        TRUNCATE TABLE dbo.ObjectsNR;
                        TRUNCATE TABLE dbo.FirstUnitCodeDevicesNR;                         
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
                      DELETE FROM ObjectsNR WHERE ProductID = {0};                      
                      DELETE FROM ProductsNR WHERE ID = {0};
                      DELETE FROM ProductsLogNR WHERE ProductID = {0};
                      DELETE FROM FirstUnitCodeDevicesNR WHERE ProductID = {0};
                      COMMIT TRANSACTION;", productId);
                SqlCommand deleteProductStructureCommand = new SqlCommand(deleteProductStructureQuery, connection);
                deleteProductStructureCommand.CommandTimeout = 0;
                deleteProductStructureCommand.ExecuteNonQuery();
            }
        }
    }
}
