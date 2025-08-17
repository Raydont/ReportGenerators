using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;

namespace OrderStructureReport
{
    internal static class ApiDocs
    {
        public static SqlConnection GetConnection(bool open)
        {
            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            var info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            var reference = info.CreateReference();
            var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
            sqlConStringBuilder.DataSource = nameSqlServer;
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

    public class DatabaseReader
    {
        public List<DocumentParameters> getDBData(int baseDocID)
        {
            SqlConnection connection = ApiDocs.GetConnection(true);
            List<DocumentParameters> docParams = new List<DocumentParameters>();

            #region Запрос, раскрывающий дерево выбранного объекта
            string getOrderTree = String.Format(
                              @"DECLARE @docid INT
                                        DECLARE @productId INT
                                        DECLARE @level INT
                                        DECLARE @insertcount INT
                                        DECLARE @CollectionID INT
                                        SET @docid = {0}
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
                                                              BomSection NVARCHAR(255),
                                                              UnitCode NVARCHAR(255) NULL) 
                                        ELSE TRUNCATE TABLE #TmpR

                                        INSERT INTO #TmpR
                                        SELECT DISTINCT n.s_ObjectID, -1, 0, n.s_ClassID,
                                                n.Name, n.Denotation, 1, 
                                                N'Комплексы', duc.Code
                                        FROM Nomenclature n INNER JOIN Documents d ON (d.s_NomenclatureObjectID = n.s_ObjectID)
                                                            LEFT JOIN UnitCodes duc ON (n.s_ObjectID = duc.s_ObjectID)                                       
                                        WHERE n.s_ObjectID=@docid AND
                                              n.s_Deleted=0 AND n.s_ActualVersion=1 AND
                                              d.s_Deleted = 0 AND d.s_ActualVersion = 1 AND    
                                              duc.s_Deleted = 0 AND duc.s_ActualVersion = 1 
      
                                        WHILE 1=1
                                        BEGIN

                                           INSERT INTO #TmpR
                                           SELECT DISTINCT nh.s_ObjectID, nh.s_ParentID, @level+1,
                                                  n.s_ClassID, n.Name, n.Denotation,
                                                  nh.Amount, nh.BOMSection, ISNULL(duc.Code,'') as UnitCode
                                           FROM NomenclatureHierarchy nh INNER JOIN #TmpR ON nh.s_ParentID = #TmpR.s_ObjectID AND
                                                                                               #TmpR.[level] = @level AND
                                                                                               nh.s_ActualVersion = 1 AND nh.s_Deleted = 0
                                                                         INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID AND
                                                                                                      n.s_ActualVersion = 1 AND n.s_Deleted = 0 AND
                                                                                                      n.s_ClassID IN (500, 501, 566)
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

                                        SELECT * FROM #TmpR", baseDocID);
            #endregion
            SqlCommand command = new SqlCommand(getOrderTree, connection);
            command.CommandTimeout = 0;
            SqlDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                var item = new DocumentParameters();
                item.ObjectID = Convert.ToInt32(reader[0]);
                item.ParentID = Convert.ToInt32(reader[1]);
                item.Level = Convert.ToInt32(reader[2]);
                item.Name = Convert.ToString(reader[4]);
                item.Denotation = reader[5] == null ? String.Empty : Convert.ToString(reader[5]);
                item.Amount = Convert.ToInt32(reader[6]);
                item.BomSection = Convert.ToString(reader[7]);
                item.UnitCode = reader[8] == null ? String.Empty : Convert.ToString(reader[8]);
                docParams.Add(item);
            }

            connection.Close();
            // Расстановка данных по родительским объектам и уровню вложенности (конструирование дерева)
            List<DocumentParameters> originalTree = new List<DocumentParameters>();
            docParams[0].pointNumber = "I";
            docParams[0].Amount = 1;
            docParams[0].TotalAmount = 1;
            FillChildren(docParams[0], originalTree, docParams);
            return originalTree;
        }

        // Пересортировка данных по родительским объектам и уровню вложенности
        private void FillChildren(DocumentParameters root, List<DocumentParameters> originalTree, List<DocumentParameters> treeData)
        {
            originalTree.Add(root);
            var children = treeData.Where(t => t.Level == root.Level + 1 && t.ParentID == root.ObjectID).Select(t => t.Clone()).ToList();

            int cnt = 0;
            foreach (var child in children)
            {
                cnt++;
                // Номер пункта структуры
                child.pointNumber = child.Level == 1 ? cnt.ToString() : root.pointNumber + "." + cnt.ToString();
                // Общее количество устройства в данной ветке
                child.TotalAmount = root.TotalAmount * child.Amount;
                FillChildren(child, originalTree, treeData);
            }
        }

        public ComplexParameters getComplexParameters(int baseDocumentID)
        {
            ComplexParameters parameters = new ComplexParameters();
            Guid nameGuid = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
            Guid denotationGuid = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
            Guid codeGuid = new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6");

            // Находим объект по его ID
            NomenclatureReference nom = new NomenclatureReference(ServerGateway.Connection);
            ReferenceObject nObject = nom.Find(baseDocumentID);

            if (nObject != null)
            {
                parameters.Name = nObject[nameGuid].Value == null ? String.Empty : nObject[nameGuid].Value.ToString();
                parameters.Denotation = nObject[denotationGuid].Value == null ? String.Empty : nObject[denotationGuid].Value.ToString();
                parameters.Order = nObject[codeGuid].Value == null ? String.Empty : nObject[codeGuid].Value.ToString();
            }
            else
            {
                parameters.Name = String.Empty;
                parameters.Denotation = String.Empty;
                parameters.Order = String.Empty;
            }

            return parameters;
        }
    }

    public class ComplexParameters
    {
        public string Name;
        public string Denotation;
        public string Order;
    }

    public class DocumentParameters
    {
        public string pointNumber;                 // Номер пункта структуры (присваивается в процессе сортировки данных)
        public int ObjectID;                       // ID объекта в Номенклатуре
        public int ParentID;                       // ID родительского объекта
        public int Level;                          // Уровень вложенности в дереве заказа
        public string Name;                        // Наименование
        public string Denotation;                  // Обозначение
        public string BomSection;                  // Раздел спецификации
        public string UnitCode;                    // Шифр устройства
        public double TotalAmount;                 // Общее количество устройства в данной ветке
        public int Amount;

        internal DocumentParameters Clone()
        {
            return new DocumentParameters() {
                pointNumber = pointNumber,          // Номер пункта структуры (присваивается в процессе сортировки данных)
                ObjectID = ObjectID,                // ID объекта в Номенклатуре
                ParentID = ParentID,                // ID родительского объекта
                Level = Level,                      // Уровень вложенности в дереве заказа
                Name = Name,                        // Наименование
                Denotation = Denotation,            // Обозначение
                BomSection = BomSection,            // Раздел спецификации
                UnitCode = UnitCode,                // Шифр устройства
                TotalAmount = TotalAmount,          // Общее количество устройства в данной ветке
                Amount = Amount,          
            };
        }
    }
}
