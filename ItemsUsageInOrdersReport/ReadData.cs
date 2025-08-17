using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;

namespace ItemsUsageInOrdersReport
{
    // типы объектов Номенклатуры в DOCs
    public static class TFDClass
    {
        public static readonly int Assembly = new NomenclatureReference().Classes.AllClasses.Find("Сборочная единица").Id;
        public static readonly int Detal = new NomenclatureReference().Classes.AllClasses.Find("Деталь").Id;
        public static readonly int Document = new NomenclatureReference().Classes.AllClasses.Find("Документ").Id;
        public static readonly int Komplekt = new NomenclatureReference().Classes.AllClasses.Find("Комплект").Id;
        public static readonly int Complex = new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Id;
        public static readonly int OtherItem = new NomenclatureReference().Classes.AllClasses.Find("Прочее изделие").Id;
        public static readonly int StandartItem = new NomenclatureReference().Classes.AllClasses.Find("Стандартное изделие").Id;
        public static readonly int Material = new NomenclatureReference().Classes.AllClasses.Find("Материал").Id;
        public static readonly int ComponentProgram = new NomenclatureReference().Classes.AllClasses.Find("Компонент (программы)").Id;
        public static readonly int ComplexProgram = new NomenclatureReference().Classes.AllClasses.Find("Комплекс (программы)").Id;
    }

    // параметры элемента
    public class TFDDocument
    {
        /* Порядок полей в запросе чтения документа */
        public int Id;
        public int ChildDocId;
        public int Class;
        public double Amount;
        public string Naimenovanie;
        public string Oboznachenie;
        public int Level;

        public TFDDocument(int id, int docClass)
        {
            Id = id;
            Class = docClass;
        }

        public override string ToString()
        {
            return Naimenovanie + " " + Oboznachenie;
        }
    }

    // данные об элементе
    public class ItemOrderBranches
    {
        int itemId;

        public ItemOrderBranches(int id)
        {
            itemId = id;
        }

        // чтение всех веток элемента
        private List<TFDDocument> ReadBranches(SqlConnection conn)
        {
            List<TFDDocument> allBrachesForItem = new List<TFDDocument>();

            SqlCommand getItemUsageCommand = new SqlCommand(String.Format(@"DECLARE @ObjID INT
                                                                                DECLARE @RCount INT
                                                                                DECLARE @Count INT

                                                                                IF OBJECT_ID('tempdb..#Input') is NULL
                                                                                CREATE TABLE #Input (s_ObjectID int, s_ChildID int, [level] int, Amount float)
                                                                                ELSE
                                                                                DELETE FROM #Input
                
                                                                                IF OBJECT_ID('tempdb..#Temp') is NULL
                                                                                CREATE TABLE #Temp (s_ObjectID int, s_ChildID int, [level] int, Amount float)
                                                                                ELSE
                                                                                DELETE FROM #Temp

                                                                                IF OBJECT_ID('tempdb..#Output') is NULL
                                                                                CREATE TABLE #Output (s_ObjectID int, s_ChildID int, [level] int, Amount float)
                                                                                ELSE
                                                                                DELETE FROM #Output                      
          
                                                                                SET @ObjID = {0}
                                                                                SET @Count = 0

                                                                                INSERT INTO #Input SELECT @ObjID,0,@Count,1
                                                                                INSERT INTO #Output SELECT @ObjID,0,@Count,1

                                                                                WHILE 1=1
                                                                                BEGIN
                                                                                   SET @Count = @Count + 1
   
                                                                                   INSERT INTO #Temp
                                                                                   SELECT h.s_ParentID,h.s_ObjectID,@Count,h.Amount
                                                                                   FROM NomenclatureHierarchy h
                                                                                   WHERE h.s_ObjectID IN (SELECT #Input.s_ObjectID FROM #Input) AND
                                                                                   h.s_ActualVersion = 1 AND h.s_Deleted = 0
   
                                                                                   INSERT INTO #Output
                                                                                   SELECT * FROM #Temp
   
                                                                                   DELETE FROM #Input
   
                                                                                   INSERT INTO #Input
                                                                                   SELECT * FROM #Temp
   
                                                                                   DELETE FROM #Temp
                                                                                   SET @RCount = (SELECT COUNT(*) FROM #Input) 
   
                                                                                   IF @RCount = 0 GOTO END1
                                                                                END
                                                                                END1:

                                                                                SELECT #Output.s_ObjectID,n.s_ClassID,n.Name,n.Denotation,#Output.s_ChildID,#Output.[level],#Output.Amount
                                                                                FROM #Output JOIN Nomenclature n ON (#Output.s_ObjectID = n.s_ObjectID)
                                                                                WHERE n.s_ActualVersion = 1 AND n.s_Deleted = 0
                                                                                ORDER BY #Output.[level],#Output.s_ObjectID", itemId), conn);
            getItemUsageCommand.CommandTimeout = 0;
            SqlDataReader reader = getItemUsageCommand.ExecuteReader();

            //SpecListParams pars = new SpecListParams(documentId);

            while (reader.Read())
            {
                int objectID = reader.GetInt32(0);
                int objectClass = GetSqlInt32(reader, 1);
                TFDDocument doc = new TFDDocument(objectID, objectClass);
                doc.Naimenovanie = GetSqlString(reader, 2);
                doc.Oboznachenie = GetSqlString(reader, 3);
                doc.ChildDocId = reader.GetInt32(4);
                doc.Level = reader.GetInt32(5);
                doc.Amount = reader.GetDouble(6);

                // добавление всех элементов в коллекцию
                allBrachesForItem.Add(doc);
                //logDelegate(String.Format("Добавлен элемент: {0} {1}", .Oboznachenie, pars.Naimenovanie));
            }
            reader.Close();

            return allBrachesForItem;
        }

        void GetAllBrunches(List<TFDDocument> current, List<TFDDocument> all, List<List<TFDDocument>> branches)
        {
            var last = current[current.Count - 1];
            if (last.Class == 502)
            {
                branches.Add(current);
                return;
            }
            var childs = all.Where(t => t.Level == last.Level + 1 && t.ChildDocId == last.Id).ToList(); ;
            if (childs.Count == 0) // не комплекс 
            {
                return;
            }

            foreach (var child in childs)
            {
                var newBrunch = current.ToList();
                newBrunch.Add(child);
                GetAllBrunches(newBrunch, all, branches);
            }

        }

        // отбор веток элемента по заказам
        public List<List<TFDDocument>> SelectionOrderBranch(SqlConnection conn, bool removeMiddle)
        {
            var itemBranches = ReadBranches(conn);

            if (itemBranches.Count == 1)
            {
                return new List<List<TFDDocument>>();
            }

            var complexBranches = new List<List<TFDDocument>>();

            Dictionary<string, int> parents = new Dictionary<string, int>();
            var brunches = new List<List<TFDDocument>>();
            GetAllBrunches(new List<TFDDocument> { itemBranches[0] }, itemBranches, brunches);

            for (int i = 0; i < brunches.Count; i++)
            {
                var brunch = brunches[i];

                var itemParent = brunch[1];

                var key = brunch[brunch.Count - 1].Id + "-" + itemParent.Id;

                if (!parents.ContainsKey(key))
                {
                    parents[key] = 0;
                    complexBranches.Add(brunch);
                }
            }

            if (removeMiddle)
            {
                complexBranches = complexBranches.Select(t => new List<TFDDocument> { t[t.Count - 1], t[1] }).ToList();
            }

            return complexBranches.OrderBy(t => t[0].Naimenovanie).ToList();
        }


        private int GetSqlInt32(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0;

            return reader.GetInt32(field);
        }

        private string GetSqlString(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return string.Empty;

            return reader.GetString(field);
        }
    }
}
