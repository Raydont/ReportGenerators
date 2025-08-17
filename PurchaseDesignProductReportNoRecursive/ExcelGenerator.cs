using System;
using System.Collections.Generic;
using System.Windows.Forms;
using TFlex.Reporting;
using PurchaseDesignProductReportForm;
using Globus.DOCs.Xlsx;

namespace PurchaseDesignProductReportNoRecursive
{
    public class ExcelGenerator : IReportGenerator
    {
        /// <summary>
        /// Расширение файла отчета
        /// </summary>
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xlsx";
            }
        }

        /// <summary>
        /// Определяет редактор параметров отчета ("Параметры шаблона" в контекстном меню Отчета)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="context"></param>
        /// <param name="owner"></param>
        /// <returns></returns>
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            context.CopyTemplateFile();    // Создаем копию шаблона
            System.IO.File.Delete(context.ReportFilePath);
            var userFullName = context.UserName;
            var m_form = new ComplexesForm();
            m_form.userFullName = userFullName;
            m_form.Visible = true;
            m_form.DotsEntry(context.ReportFilePath);
        }
    }

    
    
    // Статусы заказов в таблице ProductsLog
    public enum ProductStatus
    {
        PRODUCT_STATUS_PRODUCT_OBJECT_CREATED = 0,
        PRODUCT_STATUS_STRUCTURE_FORMING = 1,
        PRODUCT_STATUS_STRUCTURE_FORMED = 2,
        PRODUCT_STATUS_CALCULATING = 3,
        PRODUCT_STATUS_CALCULATED = 4
    }

    // Интерфейсные элементы и отчет в Excel
    public class PurchaseProductReport
    {
        XlsxDocument xls;

        // инициализация DataGridView-ов
        public void InitializeControls(DataGridView products, DataGridView addedProducts)
        {
            // столбцы DataGridView всех заказов
            DataGridViewTextBoxColumn idColumn = new DataGridViewTextBoxColumn();
            products.Columns.Add(idColumn);
            idColumn.HeaderText = "ID";
            DataGridViewTextBoxColumn nameColumn = new DataGridViewTextBoxColumn();
            products.Columns.Add(nameColumn);
            nameColumn.HeaderText = "Наименование";
            DataGridViewTextBoxColumn denotationColumn = new DataGridViewTextBoxColumn();
            products.Columns.Add(denotationColumn);
            denotationColumn.HeaderText = "Обозначение";
            DataGridViewTextBoxColumn unitCodeColumn = new DataGridViewTextBoxColumn();
            products.Columns.Add(unitCodeColumn);
            unitCodeColumn.HeaderText = "Шифр";
            DataGridViewTextBoxColumn classIdColumn = new DataGridViewTextBoxColumn();
            products.Columns.Add(classIdColumn);
            classIdColumn.HeaderText = "Тип";

            addedProducts.Columns.Clear();
            // столбцы DataGridView добавленных заказов
            DataGridViewCheckBoxColumn isSelectAddedProductsColumn = new DataGridViewCheckBoxColumn();
            isSelectAddedProductsColumn.HeaderText = "Выбран для отчета";  
            addedProducts.Columns.Add(isSelectAddedProductsColumn);                    
            DataGridViewTextBoxColumn productIdAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(productIdAddedProductsColumn);
            productIdAddedProductsColumn.HeaderText = "ProductID";
            DataGridViewTextBoxColumn objectIdAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(objectIdAddedProductsColumn);
            objectIdAddedProductsColumn.HeaderText = "ObjectID";
            DataGridViewTextBoxColumn unitCodeAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(unitCodeAddedProductsColumn);
            unitCodeAddedProductsColumn.HeaderText = "Шифр";
            DataGridViewTextBoxColumn nameAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(nameAddedProductsColumn);
            nameAddedProductsColumn.HeaderText = "Наименование";
            DataGridViewTextBoxColumn denotationAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(denotationAddedProductsColumn);
            denotationAddedProductsColumn.HeaderText = "Обозначение";
            DataGridViewTextBoxColumn dateAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(dateAddedProductsColumn);
            dateAddedProductsColumn.HeaderText = "Дата";
            DataGridViewTextBoxColumn countAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(countAddedProductsColumn);
            countAddedProductsColumn.HeaderText = "Количество изделий";
            DataGridViewTextBoxColumn statusAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(statusAddedProductsColumn);
            statusAddedProductsColumn.HeaderText = "Статус";
        }

        // заполнение DataGridView информацией о комплексах
        public void FillProducts(List<ItemInfo> complexes, DataGridView dataGridView)
        {
            dataGridView.Rows.Clear();

            foreach (ItemInfo complex in complexes)
                dataGridView.Rows.Add(complex.Id, complex.Name, complex.Denotation, complex.UnitCode, complex.ClassId);

            dataGridView.AutoResizeColumns();
        }

        // заполнение DataGridView информацией о всех добавленных комплексах
        public void FillAddedProducts(List<ComplexInfo> products, DataGridView dataGridView)
        {
            dataGridView.Rows.Clear();
            int i = 0;
            foreach (ComplexInfo product in products)
            {
                dataGridView.Rows.Add(false, product.ProductId, product.ObjectId, product.UnitCode, product.Name, product.Denotation,
                    product.CreationDate, product.CountItems, product.Comment);
                i++;
            }

            dataGridView.Refresh();
            dataGridView.AutoResizeColumns();            
        }

        // добавление комплекса в DataGridView
        public void AddProduct(ComplexInfo product, DataGridView dataGridView)
        {
            dataGridView.Rows.Add(false, product.ProductId, product.ObjectId, product.UnitCode, product.Name, product.Denotation,
                product.CreationDate, product.CountItems, product.Comment);

            dataGridView.AutoResizeColumns();
        }

        // заголовок отчета Excel
        public int GetExcelReportHeader(List<int> productIds, bool isPKIReport, bool isOccurence, bool isFirstUnitCodeDevices)
        {
            string occurenceRemark = string.Empty;
            string firstUnitCodeDevicesRemark = string.Empty;

            if (isPKIReport)
            {
                xls[3, 1].SetValue("Ведомость покупных конструкторских изделий");
                xls[3, 1].Font.Bold = true;
                xls[3, 1].CenterText();
            }
            else
            {
                if (isOccurence) occurenceRemark = " (cо списком первых входимостей)";
                if (isFirstUnitCodeDevices) firstUnitCodeDevicesRemark = " (со списком первых шифрованных устройств)";
                xls[3, 1].SetValue("Ведомость группового запуска" + occurenceRemark + firstUnitCodeDevicesRemark);
                xls[3, 1].Font.Bold = true;
                xls[3, 1].CenterText();
            }

            xls[1, 3].SetValue("Перечень заказов");
            xls[1, 3].Font.Underline = true;

            xls[1, 4].SetValue("Шифр");
            xls[2, 4].SetValue("Обозначение");
            xls[3, 4].SetValue("Наименование");
            xls[1, 4, 3, 1].Font.Bold = true;
            xls[1, 4, 3, 1].CenterText();
            xls[1, 4, 3, 1].BorderTable();

            // таблица выводимых заказов ведомости
            List<ComplexInfo> productInfos = new List<ComplexInfo>();
            int row = 5;
            for (int i = 0; i < productIds.Count; i++)
            {
                productInfos.Add(PurchaseDesignProductOperation.GetProductInfo(productIds[i]));
                xls[1, row].SetValue(productInfos[i].UnitCode);
                xls[2, row].SetValue(productInfos[i].Denotation);
                xls[3, row].SetValue(productInfos[i].Name);
                xls[1, row, 3, 1].BorderTable();
                row += 1;
            }

            row+=2;

            xls[1, row].SetValue("ID");
            xls[2, row].SetValue("Обозначение");
            xls[3, row].SetValue("Наименование");
            xls[1, row, 3, 1].Font.Bold = true;
            xls[1, row, 3, 1].CenterText();
            xls[1, row, 3, 1].BorderTable(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Medium);

            // заголовки заказов
            int col = 4;
            int mergeOffset = 0;
            for (int i = 0; i < productInfos.Count; i++)
            {
                if (isFirstUnitCodeDevices && !isOccurence)
                {
                    col = 4 + i;
                    mergeOffset = 0;
                }
                if (isOccurence && !isFirstUnitCodeDevices)
                {
                    col = 4 + 2 * i;
                    mergeOffset = 1;
                }
                if (isOccurence && isFirstUnitCodeDevices)
                {
                    col = 4 + 3 * i;
                    mergeOffset = 2;
                }
                xls[col, row].SetValue(productInfos[i].UnitCode);
                xls[col, row].Columns.AutoFit();
                xls[col, row, mergeOffset + 1, 1].Merge();
                xls[col, row].Font.Bold = true;
                xls[col, row].CenterText();
                xls[col, row, mergeOffset + 1, 1].BorderTable(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Medium);
            }

            if (isFirstUnitCodeDevices && !isOccurence) col += 1;
            if (isOccurence && !isFirstUnitCodeDevices) col += 2;
            if (isFirstUnitCodeDevices && isOccurence) col += 3;

            xls[col, row].SetValue("Общее количество");
            xls[col, row].Font.Bold = true;
            xls[col, row].CenterText();
            xls[col, row].BorderTable(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Medium);

            //xls[1, row + 1, col, 1].Select();
            xls.Worksheet.FreezePanes = true;
            xls.Worksheet.SplitRow = row;
            //xls.Application.ActiveWindow.FreezePanes = true;
            
            return ++row;
        }

        // получение отчета Excel
        public void GetExcelReportBody(List<int> productIds, bool isPKIReport, bool isOccurence, bool isFirstUnitCodeDevices, int row)
        {
            List<ItemInfo> occurence;
            List<ItemInfo> firstUnitCodeDevices;
            ItemInfo occurenceItemInfo;
            int maxOccurenceRow = 0;      // счетчик строк максимального количества вхождений для элемента в заказах
            int type;                     // тип элементов номенклатуры
            int colCount = 1;                 // количество столбцов на блок информации об элементе в заказе
            int initFirstUnitCodeCol = 6; // начальная позиция столбца первых шифрованных устройств

            // определение типа элементов отчета в зависимости от типа отчета
            if (isPKIReport) 
                type = (int)ItemType.TYPE_PKI;
            else
                type = (int)ItemType.TYPE_DETAIL;

            // определение количества столбцов на блок информации об элементе в заказе
            if (isFirstUnitCodeDevices && !isOccurence)
            {
                colCount = 1;
                initFirstUnitCodeCol = 4;
            }

            if (isOccurence && !isFirstUnitCodeDevices)
            {
                colCount = 2;
                initFirstUnitCodeCol = 4;
            }

            if (isFirstUnitCodeDevices && isOccurence)
            {
                colCount = 3;
                initFirstUnitCodeCol = 6;
            }

            // получение номенклатуры объектов отчета
            List<ItemInfo> items = PurchaseDesignProductOperation.GetReportData(productIds, type);

            foreach (ItemInfo item in items)
            {
                int row1 = row;

                // вывод Id, обозначения, наименования
                xls[1, row].SetValue(item.Id);
                xls[2, row].SetValue(item.Denotation);
                xls[3, row].SetValue(item.Name);

                row++;

                maxOccurenceRow = 0;

                using (var connection = PurchaseDesignProductOperation.GetConnectionPKI())
                {
                    connection.Open();

                    // цикл по столбцам заказов для вхождений элемента
                    for (int i = 0; i < productIds.Count; i++)
                    {
                        // общее количество элемента номенклатуры в заказе
                        int countInProductInt = PurchaseDesignProductOperation.GetObjectTotalCountInProduct(connection, item.Id, productIds[i]);
                        string countInProductString = (countInProductInt == 0 ? string.Empty : countInProductInt.ToString());
                        xls[4 + colCount * i, row - 1].SetValue(countInProductString);
                        xls[4 + colCount * i, row - 1].Font.Bold = true;
                        xls[4 + colCount * i, row - 1].CenterText();

                        if (isOccurence)
                        {
                            // получение списка вхождений (с количествами) для элемента номенклатуры
                            occurence = PurchaseDesignProductOperation.GetObjectOccurenceParams(connection, item.Id, productIds[i]);

                            // вывод списка вхождений для элемента номенклатуры в отчет
                            int occurenceRow = row;
                            for (int j = 0; j < occurence.Count; j++) // цикл по столбцу заказа
                            {
                                // получение параметров (наименование, обозначение и количество) элементов-родителей их вывод
                                occurenceItemInfo = PurchaseDesignProductOperation.GetOccurenceParams(connection, occurence[j].ParentId, productIds[i]);
                                xls[4 + colCount * i, occurenceRow].SetValue(occurenceItemInfo.Name + " (" + occurenceItemInfo.TotalCount + ")");
                                xls[4 + colCount * i, occurenceRow].NoteText = occurenceItemInfo.Denotation;
                                // количество элемента номенклатуры во вхождении
                                xls[5 + colCount * i, occurenceRow].SetValue(occurence[j].Amount);

                                if (occurence.Count > 1 && j != occurence.Count - 1) occurenceRow++;
                            }

                            if (maxOccurenceRow < occurenceRow) maxOccurenceRow = occurenceRow;
                        }

                        // вывод списка первых шифрованных устройств для элемента номенклатуры
                        if (isFirstUnitCodeDevices)
                        {
                            // получение списка первых шифрованных устройств для элемента номенклатуры
                            firstUnitCodeDevices = PurchaseDesignProductOperation.GetFirstUnitCodeDevicesForDetail(connection, item.Id, productIds[i]);

                            int unitCodeListRow = row;
                            for (int j = 0; j < firstUnitCodeDevices.Count; j++)
                            {
                                xls[initFirstUnitCodeCol + colCount * i, unitCodeListRow].SetValue(firstUnitCodeDevices[j].UnitCode + " (" + firstUnitCodeDevices[j].Amount.ToString() + ")");
                                xls[initFirstUnitCodeCol + colCount * i, unitCodeListRow].NoteText = firstUnitCodeDevices[j].Denotation;

                                if (firstUnitCodeDevices.Count > 1 && j != firstUnitCodeDevices.Count - 1) unitCodeListRow++;
                            }

                            if (maxOccurenceRow < unitCodeListRow) maxOccurenceRow = unitCodeListRow;
                        }
                    }
                }
                // общее количество элемента номенклатуры во всех заказах
                xls[4 + colCount * productIds.Count, row - 1].SetValue(item.TotalCount);
                xls[4 + colCount * productIds.Count, row - 1].CenterText();

                row = maxOccurenceRow++;
                xls[1, row1, 4 + colCount * productIds.Count, row - row1 + 1].SetBorders( DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, 
                    XlsxBorder.Top,
                    XlsxBorder.Bottom,
                    XlsxBorder.Left,
                    XlsxBorder.Right,
                    XlsxBorder.InsideVertical);
                row++;
            }
        }

        public void GetExcelReport(List<int> productIds, bool isPKIReport, bool isOccurence, bool isFirstUnitCodeDevices, string reportFilePath)
        {
            bool isCatch = false;
            xls = new XlsxDocument();
            xls.fontName = "Alial";
            xls.fontSize = 10;

            try
            {

                // заголовок отчета
                int row = GetExcelReportHeader(productIds, isPKIReport, isOccurence, isFirstUnitCodeDevices);
                // тело отчета
                GetExcelReportBody(productIds, isPKIReport, isOccurence, isFirstUnitCodeDevices, row);
                xls.AutoWidth();
                xls.SaveTo(reportFilePath);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (!isCatch)
                {
                    // Открытие файла
                    System.Diagnostics.Process.Start(reportFilePath);
                }
            }
        }
    }    
}
    



