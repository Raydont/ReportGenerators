using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TPPDataGenerator
{
    public static class Settings
    {
        // Путь до файла шаблона
        public static readonly string путьКШаблону = @"Служебные файлы\Шаблоны отчётов\Отчеты Excel\Excel.xlsx";
        // Regex для выделения документов требуемых типов
        public static readonly string требуемыеДокументыRegex = @"^(?<prefix>.*)(?<![А-Я])(Э3|СБ|ТУ|МЧ|МЭ)";
        // Regex для выделения документов-деталей
        public static readonly string всеДокументыRegex = @"^(?<prefix>.*)(?<![А-Я])(Э3|СБ|ТУ|МЧ|МЭ|РЭ|ФО|С4|ТС4|ВЭ|ПЭ4|Э0|Э4|ТЭ4|УП)";
        // Regex для выделения объектов ЕСПД
        public static readonly string ЕСПДRegex = @"^(?:[А-Я]{4})\.(?:[0-9]{5})-(?:[0-9]{2})(?:\x20(?:[0-9]{2}))?(?:\x20(?:[0-9]{2}))?(?:-(?:[0-9]))?(?:-(?:ЛУ)|-(?:УД)|-(?:УЛ))?$";
        // Сдвиг по списку объектов (для обработки порциями)
        public static int сдвиг = 100;
        // Количество столбцов в отчете (файл Excel)
        public static int столбцыКоличество = 7;

        public static class Запросы
        {
            /* Запрос, определяющий, в какие ЭСЗ входят выбранные объекты Номенклатуры
               Используется для нового КД или ИИ
               Параметры запроса:
               0 - типы объектов ЭСЗ - родителей (структура, описание, ДСЕ для разработки, вариант ЭСИ)
               1 - типы объектов ЭСЗ - не включаемые в список для проверки (описание, ДСЕ для разработки, документ, прочее и т.п.)
               2 - тип объекта ЭСЗ - тип объекта Вариант ЭСИ
               3 - список ID выбранных объектов номенклатуры
               4 - тип объекта Номенклатуры - Документ
               5 - тип объекта Номенклатуры - Деталь */
            public static readonly string объектыЭСИвЭСЗ =
              @"DECLARE @docID INT;
                DECLARE @level INT; SET @level = 0;
                DECLARE @insertcount INT; SET @insertcount = 0;

                IF OBJECT_ID('tempdb..#StructureTree') IS NULL
                CREATE TABLE #StructureTree (s_StructID INT,
                                         s_ObjectID INT,   
										 s_ParentID INT,										 
                                         level INT,
                                         s_ClassID INT)
                ELSE TRUNCATE TABLE #StructureTree

                -- Деревья объектов справочника ЭСЗ в составе списка актуальных (незакрытых) ЭСЗ
                INSERT INTO #StructureTree
                SELECT DISTINCT n.s_ObjectID, 
                                n.s_ObjectID, 
                                0, 
								0,
                                n.s_ClassID
                FROM glob_OrderStructureRef n
                WHERE n.s_ObjectID IN (SELECT os.s_ObjectID
                                       FROM glob_OrderStructureRef os
                                            INNER JOIN glob_OrderNumLink onl ON onl.MasterID = os.s_ObjectID
                                            INNER JOIN glob_OrderNumber num ON num.s_ObjectID = onl.SlaveID
                                       WHERE num.OrderClosed = 0)
                      AND n.s_Deleted = 0 
                      AND n.s_ActualVersion = 1
    
                WHILE 1=1
                BEGIN
                    INSERT INTO #StructureTree
                    SELECT DISTINCT st.s_StructID, 
                                    os.s_ObjectID, 
									os.s_ParentID,
                                    @level+1, 
                                    os.s_ClassID
                    FROM glob_OrderStructureRef os 
                         INNER JOIN #StructureTree st ON os.s_ParentID = st.s_ObjectID 
                                               AND st.level = @level 
                                               AND os.s_ActualVersion = 1 
                                               AND os.s_Deleted = 0
                    SET @insertcount = @@ROWCOUNT 
                    SET @level = @level + 1  

                    IF  @insertcount = 0
                    GOTO END1
                END             	
                END1:

				-- Все объекты-родители в дереве ЭСЗ нужных типов
				SELECT * 
				INTO #Parents
				FROM #StructureTree st
				WHERE st.s_ObjectID IN	(SELECT DISTINCT st1.s_ParentID
				                         FROM #StructureTree st1)
                      AND st.s_ClassID IN ({0})
				ORDER BY s_ParentID, s_ObjectID

				-- Все объекты ЭСЗ, родители которых входят таблицу Parents исключая определенные типы
				SELECT * 
				INTO #FilteredStructureTree
				FROM #StructureTree st
				WHERE st.s_ParentID IN (SELECT p.s_ObjectID FROM #Parents p)
				AND st.s_ClassID NOT IN ({1})

                -- Связь объектов ЭСЗ с объектами ЭСИ (номенклатуры)
                SELECT ft.s_StructID AS s_StructID, 
				       ft.s_ObjectID,
					   ft.s_ParentID,
				       ft.s_ClassID AS s_ClassID,
                       nl.SlaveID AS s_NomID
                INTO #StrToNomenclatureLink
                FROM #FilteredStructureTree ft
                     INNER JOIN glob_OSNomenclatureLink nl ON nl.MasterID = ft.s_ObjectID

				-- Объекты - варианты номенклатуры
                SELECT * 
				INTO #NomenclatureVariants
				FROM #StrToNomenclatureLink sl
				WHERE sl.s_ClassID = {2}

				-- Дочерние объекты вариантов, не входящие в ЭСЗ
				SELECT nv.s_StructID, 
				       0 AS s_ObjectID, 
					   nh.s_ParentID, 
					   0 AS s_ClassID, 
					   nh.s_ObjectID As s_NomID
				INTO #VarToNomenclatureLink
				FROM NomenclatureHierarchy nh
				INNER JOIN #NomenclatureVariants nv ON (nh.s_ParentID = nv.s_NomID
				                                     AND nh.s_ActualVersion = 1
													 AND nh.s_Deleted = 0)
				INNER JOIN Nomenclature n ON (nh.s_ObjectID = n.s_ObjectID
				                              AND n.s_ActualVersion = 1 AND n.s_Deleted = 0)
				AND n.ShowInStructureDoc = 0

				-- Соединение объектов, взятых из ЭСЗ, 
				-- и дочерних объектов вариантов, не входящих в ЭСЗ
				SELECT * 
				INTO #NomenclatureTreeInitial
				FROM #StrToNomenclatureLink

				UNION

				SELECT * FROM #VarToNomenclatureLink

                -- Раскрытие дерева номенклатуры каждого из объектов ЭСИ, отобранного из ЭСЗ
                IF OBJECT_ID('tempdb..#NomenclatureTree') IS NULL
                CREATE TABLE #NomenclatureTree (s_RootObjectID INT,
                                                 s_ObjectID INT,             
                                                 level INT) 
                ELSE TRUNCATE TABLE #NomenclatureTree

                SET @level = 0;
                SET @insertcount = 0;

                -- Запись начальных условий для цикла - все объекты, отобранные из ЭСЗ (исключая варианты)
                INSERT INTO #NomenclatureTree
				SELECT nti.s_NomID,
				       nti.s_NomID,
					   0
				FROM #NomenclatureTreeInitial nti
				WHERE nti.s_ClassID <> {2}
  
                -- Цикл получения дерева каждого из объектов таблицы #NomenclatureTreeInitial
                WHILE 1=1
                BEGIN

                    INSERT INTO #NomenclatureTree
                    SELECT nt.s_RootObjectID, 
                           nh.s_ObjectID, 
                           @level+1
                    FROM NomenclatureHierarchy nh 
                         INNER JOIN #NomenclatureTree nt ON nh.s_ParentID = nt.s_ObjectID 
                                               AND nt.level = @level 
                                               AND nh.s_ActualVersion = 1 
                                               AND nh.s_Deleted = 0

                    SET @insertcount = @@ROWCOUNT 
                    SET @level = @level + 1

                    IF  @insertcount = 0
                    GOTO END2
                END             	
                END2:

				-- По окончании добавляются объекты-варианты (т.к. их дерево раскрывать не надо)
                INSERT INTO #NomenclatureTree
				SELECT nti.s_NomID,
				       nti.s_NomID,
					   0
				FROM #NomenclatureTreeInitial nti
				WHERE nti.s_ClassID = {2}

				-- Основные исполнения объектов номенклатуры
				SELECT n.s_ObjectID AS NomID, 
                       n.s_ClassID AS NomClassID, 
				       CASE WHEN nbv.SlaveID > 0
				            THEN nbv.SlaveID
				            ELSE n.s_ObjectID 
                       END AS BaseNomID
			    INTO #NomBaseVersions
				FROM Nomenclature n
				     LEFT JOIN NomenclatureBaseVersion nbv ON (n.s_ObjectID = nbv.MasterID)
				WHERE n.s_ObjectID IN ({3}) 
                      AND n.s_ActualVersion = 1 
                      AND n.s_Deleted = 0

                -- Получение извещений для объектов ЭСИ
                -- Исключаются удаленные связи с извещениями (по полю DeletedChangelistID)
                SELECT cn.CreationDate AS CreationDate, 
                       cn.Name AS Name, 
                       cna.SlaveID AS SlaveID
				INTO #NomChangeNotifications
				FROM CNApplicability cna
				     INNER JOIN ChangeNotifications cn ON (cn.s_ObjectID = cna.MasterID)
				WHERE cna.SlaveID IN ({3}) 
                      AND cn.s_ActualVersion = 1 
                      AND cn.s_Deleted = 0
                      AND cna.DeletedChangelistID = 0 
                      AND cna.AddedClientViewID = 0

                -- Получение извещений для базовых исполнений объектов ЭСИ
				SELECT cn.CreationDate AS CreationDate, 
                       cn.Name AS Name, 
                       cna.SlaveID AS SlaveID
				INTO #NomBaseChangeNotifications
				FROM CNApplicability cna
				     INNER JOIN ChangeNotifications cn ON (cn.s_ObjectID = cna.MasterID)
				WHERE cna.SlaveID IN (SELECT DISTINCT nbv.BaseNomID FROM #NomBaseVersions nbv) 
                      AND cn.s_ActualVersion = 1 
                      AND cn.s_Deleted = 0
                      AND cna.DeletedChangelistID = 0 
                      AND cna.AddedClientViewID = 0

                -- Присоединение к извещениям для объектов ЭСИ извещений для основных исполнений этих объектов
				SELECT * 
				INTO #NomAllChangeNotifications
				FROM #NomChangeNotifications

				UNION

				SELECT nbcn.CreationDate AS CreationDate,
                       nbcn.Name AS Name, 
                       nbv.NomID AS SlaveID 
                FROM #NomBaseVersions as nbv
				     LEFT JOIN #NomBaseChangeNotifications as nbcn 
				     ON nbv.BaseNomID = nbcn.SlaveID
				
                -- Получение последних по времени извещений для объектов ЭСИ
				SELECT cn1.Name AS Name, 
                       cn1.CreationDate AS CreationDate, 
                       cn1.SlaveID AS SlaveID
                INTO #LastChangeNotifications
				FROM #NomAllChangeNotifications cn1 
                     LEFT JOIN #NomAllChangeNotifications cn2 ON (cn1.SlaveID = cn2.SlaveID 
                                            AND cn1.CreationDate < cn2.CreationDate)
				WHERE cn2.CreationDate IS NULL

				-- Сборочные чертежи в составе объектов номенклатуры
				SELECT nbv.NomID AS NomID, 
                       nbv.NomClassID AS NomClassID, 
                       nbv.BaseNomID AS BaseNomID, 
                       nh.s_ObjectID AS ChildID
				INTO #AssemblyDocuments
				FROM #NomBaseVersions nbv
				     LEFT JOIN NomenclatureHierarchy nh ON (nh.s_ParentID = nbv.BaseNomID 
                                                            AND nh.s_ActualVersion = 1 
                                                            AND nh.s_Deleted = 0)
				     LEFT JOIN Nomenclature n ON (n.s_ObjectID = nh.s_ObjectID 
                                                  AND n.s_ActualVersion = 1 
                                                  AND n.s_Deleted = 0)
				WHERE n.s_ClassID = {4} AND n.Denotation LIKE '%СБ%'

				-- Соединение основных исполнений и дочерних документов
				SELECT nbv.NomID AS NomID,  				        
				        CASE WHEN nbv.NomClassID = {5}
				             THEN nbv.BaseNomID 
						     ELSE ad.ChildID 
                        END AS ChildID
				INTO #BaseVersionsToAssemblyDocsLink
				FROM #NomBaseVersions nbv
				     LEFT JOIN #AssemblyDocuments ad ON (nbv.BaseNomID = ad.BaseNomID 
				                         AND nbv.NomID = ad.NomID)

				-- Таблица с признаком наличия КД
				SELECT bvad.NomID AS NomID,
				       CASE WHEN COUNT(f.s_ObjectID) > 0 
                            THEN 'Да' 
                            ELSE 'Нет' 
                       END AS HasCD
				INTO #NomObjectHasCD
				FROM #BaseVersionsToAssemblyDocsLink bvad
				     LEFT JOIN Nomenclature n ON (bvad.ChildID = n.s_ObjectID 
                                                  AND n.s_ActualVersion = 1 
                                                  AND n.s_Deleted = 0)
				     LEFT JOIN DocumentFiles df ON (n.s_LinkedObjectID = df.MasterID)
				     LEFT JOIN Files f ON (f.s_ObjectID = df.SlaveID 
                                           AND f.s_ActualVersion = 1 
                                           AND f.s_Deleted = 0)
				GROUP BY bvad.NomID

                -- Итоговая таблица - Добавление к дереву объектов ЭСИ, входящих в состав ЭСЗ, деревьев каждого из этих объектов ЭСИ,
                -- а также добавление параметров объектов Номенклатуры, ТП, стадий, извещений
                -- Исключаются удаленные и измененные связи с техпроцессами (по полям DeletedChangelistID, CreatedChangelistID, AddedClientViewID)
                SELECT DISTINCT nti.s_StructID, 
                                sp.OrderNum, 
                                sr.Denotation, 
                                nt.s_ObjectID, 
                                n.Denotation, 
                                n.Name,
                                cl.Name, 
                                stg.Name, 
                                cn.Name, 
                                FORMAT(cn.CreationDate,'dd/MM/yyyy'), 
                                tps.primechanie, 
                                hcd.HasCD
                FROM #NomenclatureTreeInitial nti
                     INNER JOIN #NomenclatureTree nt ON (nti.s_NomID = nt.s_RootObjectID)
			         LEFT JOIN glob_OrderStructureParam sp ON (nti.s_StructID = sp.s_ObjectID)
                     LEFT JOIN glob_OrderStructureRef sr ON (nti.s_StructID = sr.s_ObjectID)
			         LEFT JOIN Nomenclature n ON (nt.s_ObjectID = n.s_ObjectID 
                                                  AND n.s_Deleted = 0 
                                                  AND n.s_ActualVersion = 1)
                     LEFT JOIN Classes cl ON (n.s_ClassID = cl.PK)
	                 LEFT JOIN TPLinksNomenclatureDSE tpl ON (tpl.SlaveID = nt.s_ObjectID 
                                                              AND tpl.AddedClientViewID = 0 
                                                              AND tpl.CreatedChangelistID = 0
                                                              AND tpl.DeletedChangelistID = 0)
					 LEFT JOIN TechnologicalProcesses tp ON (tp.s_ObjectID = tpl.MasterID 
                                                             AND tp.s_Deleted = 0 
                                                             AND tp.s_ActualVersion = 1)
                     LEFT JOIN glob_TP_Signs tps ON (tp.s_ObjectID = tps.s_ObjectID 
                                                     AND tps.s_Deleted = 0 
                                                     AND tps.s_ActualVersion = 1)
					 LEFT JOIN Stages stg ON (stg.PK = tp.s_StageID)
					 LEFT JOIN #LastChangeNotifications cn ON (cn.SlaveID = n.s_ObjectID)
					 LEFT JOIN #NomObjectHasCD hcd ON (hcd.NomID = n.s_ObjectID)
                     WHERE nt.s_ObjectID IN ({3})
                UNION
				SELECT 0, 
                       'Не входят в актуальные ЭСЗ',
                       '', 
                       n.s_ObjectID, 
                       n.Denotation, 
                       n.Name, 
                       cl.Name, 
                       st.Name, 
                       cn.Name, 
                       FORMAT(cn.CreationDate,'dd/MM/yyyy'), 
                       tps.primechanie, 
                       hcd.HasCD
				FROM Nomenclature n
			         LEFT JOIN Classes cl ON (n.s_ClassID = cl.PK)
					 LEFT JOIN TPLinksNomenclatureDSE tpl ON (tpl.SlaveID = n.s_ObjectID 
                                                              AND tpl.AddedClientViewID = 0 
                                                              AND tpl.CreatedChangelistID = 0 
                                                              AND tpl.DeletedChangelistID = 0)
					 LEFT JOIN TechnologicalProcesses tp ON (tp.s_ObjectID = tpl.MasterID 
                                                             AND tp.s_Deleted = 0 
                                                             AND tp.s_ActualVersion = 1)
                     LEFT JOIN glob_TP_Signs tps ON (tp.s_ObjectID = tps.s_ObjectID 
                                                     AND tps.s_Deleted = 0 
                                                     AND tps.s_ActualVersion = 1)
					 LEFT JOIN Stages st ON (st.PK = tp.s_StageID)
					 LEFT JOIN #LastChangeNotifications cn ON (cn.SlaveID = n.s_ObjectID)
					 LEFT JOIN #NomObjectHasCD hcd ON (hcd.NomID = n.s_ObjectID)
				     WHERE n.s_ObjectID IN ({3}) 
                           AND n.s_ActualVersion = 1 
                           AND n.s_Deleted = 0 
                           AND n.s_ObjectID NOT IN (SELECT DISTINCT nt.s_ObjectID FROM #NomenclatureTree nt)";

            /* Запрос, определяющий, у каких объектов Номенклатуры, входящих в выбранные ЭСЗ, неактуальные ТП
               Используется для ЭСЗ (или группы ЭСЗ)
               Параметры запроса:
               0 - типы объектов ЭСЗ - родителей (структура, описание, ДСЕ для разработки, вариант ЭСИ)
               1 - типы объектов ЭСЗ - не включаемые в список для проверки (описание, ДСЕ для разработки, документ, прочее и т.п.)
               2 - тип объекта ЭСЗ - Вариант ЭСИ
               3 - типы объектов Номенклатуры, подлежащие проверке
               4 - тип объекта Номенклатуры - Документ
               5 - тип объекта Номенклатуры - Деталь 
               6 - стадия ТП - ТП утвержден */
            public static readonly string объектыЭСИвВыгружаемыхЭСЗ =
              @"DECLARE @docID INT;
                DECLARE @level INT; SET @level = 0;
                DECLARE @insertcount INT; SET @insertcount = 0;

                IF OBJECT_ID('tempdb..#StructureTree') IS NULL
                CREATE TABLE #StructureTree (s_StructID INT,
                                         s_ObjectID INT, 
					                     s_ParentID INT,
                                         level INT,
                                         s_ClassID INT)
                ELSE TRUNCATE TABLE #StructureTree

                -- Дерево объектов справочника ЭСЗ в составе ЭСЗ
                INSERT INTO #StructureTree
                SELECT DISTINCT n.s_ObjectID, 
                                n.s_ObjectID, 
				                0,
                                0, 
                                n.s_ClassID
                FROM glob_OrderStructureRef n
                      INNER JOIN glob_OrderNumLink onl ON onl.MasterID = n.s_ObjectID
                      INNER JOIN glob_OrderNumber num ON num.s_ObjectID = onl.SlaveID
                WHERE n.s_ObjectID IN (##ЭСЗ##)
                      AND num.OrderClosed = 0
                      AND n.s_Deleted = 0 
                      AND n.s_ActualVersion = 1      
    
                WHILE 1=1
                BEGIN
                    INSERT INTO #StructureTree
                    SELECT DISTINCT st.s_StructID, 
                                    os.s_ObjectID, 
									os.s_ParentID,
                                    @level+1, 
                                    os.s_ClassID
                    FROM glob_OrderStructureRef os 
                    INNER JOIN #StructureTree st ON os.s_ParentID = st.s_ObjectID
                                          AND st.level = @level 
                                          AND os.s_ActualVersion = 1 
                                          AND os.s_Deleted = 0
                    SET @insertcount = @@ROWCOUNT 
                    SET @level = @level + 1  

                    IF  @insertcount = 0
                    GOTO END1
                END             	
                END1:

				-- Все объекты-родители в дереве ЭСЗ нужных типов
				SELECT * 
				INTO #Parents
				FROM #StructureTree st
				WHERE st.s_ObjectID IN	(SELECT DISTINCT st1.s_ParentID
				                         FROM #StructureTree st1)
                      AND st.s_ClassID IN ({0})
				ORDER BY s_ParentID, s_ObjectID

				-- Все объекты ЭСЗ, родители которых входят таблицу Parents исключая определенные типы
				SELECT * 
				INTO #FilteredStructureTree
				FROM #StructureTree st
				WHERE st.s_ParentID IN (SELECT p.s_ObjectID FROM #Parents p)
				AND st.s_ClassID NOT IN ({1})

                -- Связь объектов ЭСЗ с объектами ЭСИ (номенклатуры)
                SELECT ft.s_StructID AS s_StructID, 
				       ft.s_ObjectID,
					   ft.s_ParentID,
				       ft.s_ClassID AS s_ClassID,
                       nl.SlaveID AS s_NomID
                INTO #StrToNomenclatureLink
                FROM #FilteredStructureTree ft
                     INNER JOIN glob_OSNomenclatureLink nl ON nl.MasterID = ft.s_ObjectID

				-- Объекты - варианты номенклатуры
                SELECT * 
				INTO #NomenclatureVariants
				FROM #StrToNomenclatureLink sl
				WHERE sl.s_ClassID = {2}

				-- Дочерние объекты вариантов, не входящие в ЭСЗ
				SELECT nv.s_StructID, 
				       0 AS s_ObjectID, 
					   nh.s_ParentID, 
					   0 AS s_ClassID, 
					   nh.s_ObjectID As s_NomID
				INTO #VarToNomenclatureLink
				FROM NomenclatureHierarchy nh
				INNER JOIN #NomenclatureVariants nv ON (nh.s_ParentID = nv.s_NomID
				                                     AND nh.s_ActualVersion = 1
													 AND nh.s_Deleted = 0)
				INNER JOIN Nomenclature n ON (nh.s_ObjectID = n.s_ObjectID
				                              AND n.s_ActualVersion = 1 AND n.s_Deleted = 0)
				AND n.ShowInStructureDoc = 0

				-- Соединение объектов, взятых из ЭСЗ, 
				-- и дочерних объектов вариантов, не входящих в ЭСЗ           
				SELECT * 
				INTO #NomenclatureTreeInitial
				FROM #StrToNomenclatureLink

				UNION

				SELECT * FROM #VarToNomenclatureLink

                -- Раскрытие дерева номенклатуры каждого из объектов ЭСИ, отобранного из ЭСЗ
                IF OBJECT_ID('tempdb..#NomenclatureTree') IS NULL
                CREATE TABLE #NomenclatureTree (s_RootObjectID INT,
                                                 s_ObjectID INT,    
                                                 level INT) 
                ELSE TRUNCATE TABLE #NomenclatureTree

                SET @level = 0;
                SET @insertcount = 0;

                -- Запись начальных условий для цикла - все объекты, отобранные из ЭСЗ (исключая варианты)
                INSERT INTO #NomenclatureTree
				SELECT nti.s_NomID,
				       nti.s_NomID,
					   0
				FROM #NomenclatureTreeInitial nti
				WHERE nti.s_ClassID <> {2}

                -- Цикл получения дерева каждого из объектов таблицы #NomenclatureTreeInitial
                WHILE 1=1
                BEGIN

                    INSERT INTO #NomenclatureTree
                    SELECT nt.s_RootObjectID, 
                           nh.s_ObjectID, 
                           @level+1
                    FROM NomenclatureHierarchy nh 
                         INNER JOIN #NomenclatureTree nt ON nh.s_ParentID = nt.s_ObjectID 
                                               AND nt.level = @level 
                                               AND nh.s_ActualVersion = 1 
                                               AND nh.s_Deleted = 0

                    SET @insertcount = @@ROWCOUNT 
                    SET @level = @level + 1

                    IF  @insertcount = 0
                    GOTO END2
                END             	
                END2:

				-- По окончании добавляются объекты-варианты (т.к. их дерево раскрывать не надо)
                INSERT INTO #NomenclatureTree
				SELECT nti.s_NomID,
				       nti.s_NomID,
					   0
				FROM #NomenclatureTreeInitial nti
				WHERE nti.s_ClassID = {2}

                -- Добавление к дереву объектов ЭСИ, входящих в состав ЭСЗ, деревьев каждого из этих объектов ЭСИ
                SELECT DISTINCT nti.s_StructID AS StructID, 
                                sp.OrderNum AS OrderNum, 
                                sr.Denotation AS StructDenotation, 
                                nt.s_ObjectID AS NomID, 
                                n.Denotation AS NomDenotation, 
                                n.Name AS NomName, 
                                cl.Name AS NomClass, 
                                n.s_ClassID AS NomClassID, 
                                n.TPNonActual AS TPNonActual
                INTO #NomCompleteTree
                FROM #NomenclatureTreeInitial nti
                     INNER JOIN #NomenclatureTree nt ON (nti.s_NomID = nt.s_RootObjectID)
			         LEFT JOIN glob_OrderStructureParam sp ON (nti.s_StructID = sp.s_ObjectID)
                     LEFT JOIN glob_OrderStructureRef sr ON (nti.s_StructID = sr.s_ObjectID)
			         LEFT JOIN Nomenclature n ON (nt.s_ObjectID = n.s_ObjectID 
                                                  AND n.s_Deleted = 0 
                                                  AND n.s_ActualVersion = 1)
			         LEFT JOIN Classes cl ON (n.s_ClassID = cl.PK)
                WHERE n.s_ClassID IN ({3})

				-- Основные исполнения объектов номенклатуры
				SELECT nct.NomID AS NomID, 
                       nct.NomClassID AS NomClassID, 
				       CASE WHEN nbv.SlaveID > 0
				            THEN nbv.SlaveID
				            ELSE nct.NomID 
                       END AS BaseNomID
			    INTO #NomBaseVersions
				FROM #NomCompleteTree nct
				     LEFT JOIN NomenclatureBaseVersion nbv ON (nct.NomID = nbv.MasterID)

				-- Сборочные чертежи в составе объектов номенклатуры
				SELECT nbv.NomID AS NomID, 
                       nbv.NomClassID AS NomClassID, 
                       nbv.BaseNomID, nh.s_ObjectID AS ChildID
				INTO #AssemblyDocuments
				FROM #NomBaseVersions nbv
				     LEFT JOIN NomenclatureHierarchy nh ON (nh.s_ParentID = nbv.BaseNomID 
                                                            AND nh.s_ActualVersion = 1 
                                                            AND nh.s_Deleted = 0)
				     LEFT JOIN Nomenclature n ON (n.s_ObjectID = nh.s_ObjectID 
                                                  AND n.s_ActualVersion = 1 
                                                  AND n.s_Deleted = 0)
				WHERE n.s_ClassID = {4} 
                      AND n.Denotation LIKE '%СБ%'

				-- Соединение основных исполнений и дочерних документов
				SELECT nbv.NomID AS NomID,  				        
				        CASE WHEN nbv.NomClassID = {5} 
				             THEN nbv.BaseNomID 
						     ELSE ad.ChildID 
                        END AS ChildID
				INTO #BaseVersionsToAssemblyDocsLink
				FROM #NomBaseVersions nbv
				     LEFT JOIN #AssemblyDocuments ad ON (nbv.BaseNomID = ad.BaseNomID 
				                         AND nbv.NomID = ad.NomID)

				-- Таблица с признаком наличия КД
				SELECT bvad.NomID,
				       CASE WHEN COUNT(f.s_ObjectID) > 0 
                            THEN 'Да' 
                            ELSE 'Нет' 
                       END AS HasCD
				INTO #NomObjectHasCD
				FROM #BaseVersionsToAssemblyDocsLink bvad
				      LEFT JOIN Nomenclature n ON (bvad.ChildID = n.s_ObjectID 
                                                   AND n.s_ActualVersion = 1 
                                                   AND n.s_Deleted = 0)
				      LEFT JOIN DocumentFiles df ON (n.s_LinkedObjectID = df.MasterID)
				      LEFT JOIN Files f ON (f.s_ObjectID = df.SlaveID 
                                            AND f.s_ActualVersion = 1 
                                            AND f.s_Deleted = 0)
				GROUP BY bvad.NomID

                -- Получение извещений для объектов ЭСИ
                -- Исключаются удаленные связи с извещениями (по полю DeletedChangelistID)
                SELECT cn.CreationDate AS CreationDate, 
                       cn.Name AS Name, 
                       cna.SlaveID AS SlaveID
				INTO #NomChangeNotifications
				FROM CNApplicability cna
				     INNER JOIN ChangeNotifications cn ON (cn.s_ObjectID = cna.MasterID)
				WHERE cna.SlaveID IN (SELECT DISTINCT nct.NomID FROM #NomCompleteTree nct) 
                      AND cn.s_ActualVersion = 1 
                      AND cn.s_Deleted = 0
                      AND cna.DeletedChangelistID = 0 
                      AND cna.AddedClientViewID = 0

                -- Получение извещений для базовых исполнений объектов ЭСИ
				SELECT cn.CreationDate AS CreationDate, 
                       cn.Name AS Name, 
                       cna.SlaveID AS SlaveID
				INTO #NomBaseChangeNotifications
				FROM CNApplicability cna
				     INNER JOIN ChangeNotifications cn ON (cn.s_ObjectID = cna.MasterID)
				WHERE cna.SlaveID IN (SELECT DISTINCT nbv.BaseNomID FROM #NomBaseVersions nbv) 
                      AND cn.s_ActualVersion = 1 
                      AND cn.s_Deleted = 0
                      AND cna.DeletedChangelistID = 0 
                      AND cna.AddedClientViewID = 0

                -- Присоединение к извещениям для объектов ЭСИ извещений для основных исполнений этих объектов
				SELECT * 
				INTO #NomAllChangeNotifications
				FROM #NomChangeNotifications

				UNION

				SELECT nbcn.CreationDate AS CreationDate,
                       nbcn.Name AS Name, 
                       nbv.NomID AS SlaveID 
                FROM #NomBaseVersions as nbv
				     LEFT JOIN #NomBaseChangeNotifications as nbcn 
				     ON nbv.BaseNomID = nbcn.SlaveID

                -- Получение последних по времени извещений для объектов ЭСИ
				SELECT cn1.Name AS Name,
                       cn1.CreationDate AS CreationDate, 
                       cn1.SlaveID AS SlaveID
                INTO #LastChangeNotifications
				FROM #NomAllChangeNotifications cn1 
                     LEFT JOIN #NomAllChangeNotifications cn2 ON (cn1.SlaveID = cn2.SlaveID 
                                            AND cn1.CreationDate < cn2.CreationDate)
				WHERE cn2.CreationDate IS NULL

                -- Объекты ЭСИ с актуальными ТП
				SELECT DISTINCT tpl.SlaveID AS SlaveID
				INTO #NomActualTP
                FROM TPLinksNomenclatureDSE tpl
                     INNER JOIN TechnologicalProcesses tp ON (tp.s_ObjectID = tpl.MasterID)
                WHERE tpl.SlaveID IN (SELECT DISTINCT nct.NomID FROM #NomCompleteTree nct) 
                      AND tp.s_StageID = {6}

                -- Итоговая таблица - добавление параметров ТП, стадий, извещений
                -- Исключаются удаленные и измененные связи с техпроцессами (по полям DeletedChangelistID, CreatedChangelistID, AddedClientViewID)
				SELECT nct.StructID, 
                       nct.OrderNum, 
                       nct.StructDenotation, 
                       nct.NomID, 
                       nct.NomDenotation, 
                       nct.NomName, 
                       nct.NomClass, 
                       stg.Name, 
                       cn.Name, 
                       FORMAT(cn.CreationDate,'dd/MM/yyyy'), 
                       tps.primechanie, 
                       hcd.HasCD
				FROM #NomCompleteTree nct
				     LEFT JOIN TPLinksNomenclatureDSE tpl ON (tpl.SlaveID = nct.NomID AND tpl.AddedClientViewID = 0 
                                                              AND tpl.CreatedChangelistID = 0 
                                                              AND tpl.DeletedChangelistID = 0)
				     LEFT JOIN TechnologicalProcesses tp ON (tp.s_ObjectID = tpl.MasterID 
                                                             AND tp.s_Deleted = 0 
                                                             AND tp.s_ActualVersion = 1)
					 LEFT JOIN glob_TP_Signs tps ON (tp.s_ObjectID = tps.s_ObjectID 
                                                     AND tps.s_Deleted = 0 
                                                     AND tps.s_ActualVersion = 1)
				     LEFT JOIN Stages stg ON (stg.PK = tp.s_StageID)
				     LEFT JOIN #LastChangeNotifications cn ON (cn.SlaveID = nct.NomID)
                     LEFT JOIN #NomObjectHasCD hcd ON (hcd.NomID = nct.NomID)
                     WHERE nct.NomID NOT IN (SELECT ntp.SlaveID FROM #NomActualTP ntp) 
                        OR nct.TPNonActual = 1";

            public static readonly string выгружаемыеЭСЗ =
              @"SELECT os.s_ObjectID
                FROM glob_OrderStructureRef os
                WHERE os.Unload1C = 1
				      AND os.s_ActualVersion = 1
				      AND os.s_Deleted = 0";
        }
    }
}
