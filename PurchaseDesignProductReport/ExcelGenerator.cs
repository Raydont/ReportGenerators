using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References.Nomenclature;
using PurchaseDesignProductReportForm;
using TFlex.Reporting;
using ReportHelpers;
using Microsoft.Office.Interop.Excel;

namespace PurchaseDesignProductReport
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
                return ".xls";
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
            ComplexesForm m_form = new ComplexesForm();
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
        Xls xls;

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
        public int GetExcelReportHeader(List<int> productIds, bool isPKIReport)
        {
            if (isPKIReport)
            {
                xls[3, 1].SetValue("Ведомость покупных конструкторских изделий");
                xls[3, 1].Font.Bold = true;
                xls[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[3, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            }
            else
            {
                xls[3, 1].SetValue("Ведомость группового запуска");
                xls[3, 1].Font.Bold = true;
                xls[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[3, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            }

            xls[1, 3].SetValue("Перечень заказов");
            xls[1, 3].Font.Underline = true;

            xls[1, 4].SetValue("Шифр");
            xls[1, 4].Font.Bold = true;

            xls[2, 4].SetValue("Обозначение");
            xls[2, 4].Font.Bold = true;

            xls[3, 4].SetValue("Наименование");
            xls[3, 4].Font.Bold = true;

            xls[1, 4, 3, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);

            // таблица выводимых заказов ведомости
            List<ComplexInfo> productInfos = new List<ComplexInfo>();
            int row = 5;
            for (int i = 0; i < productIds.Count; i++)
            {
                row += i;
                productInfos.Add(PurchaseDesignProductOperation.GetProductInfo(productIds[i]));
                xls[1, row].SetValue(productInfos[i].UnitCode);
                xls[2, row].SetValue(productInfos[i].Denotation);
                xls[3, row].SetValue(productInfos[i].Name);
                xls[1, row, 3, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            }

            row+=2;

            xls[1, row].SetValue("ID");
            xls[2, row].SetValue("Обозначение");
            xls[3, row].SetValue("Наименование");
            xls[1, row, 3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 3, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, row, 3, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

            // заголовки заказов
            int col = 4;
            for (int i = 0; i < productInfos.Count; i++)
            {
                col = 4 + 2*i;
                xls[col, row].SetValue(productInfos[i].UnitCode);
                xls[col, row, 2, 1].Merge();
                xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                xls[col, row].Font.Bold = true;
                xls[col, row, 2, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }

            col += 2;

            xls[col, row].SetValue("Общее количество");
            xls[col, row].Font.Bold = true;
            xls[col, row].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

            xls[1, row + 1, col, 1].Select();
            xls.Application.ActiveWindow.FreezePanes = true;

            return ++row;
        }

        // получение отчета Excel
        public void GetExcelReportBody(List<int> productIds, bool isPKIReport, int row)
        {
            List<ItemInfo> occurence;
            ItemInfo occurenceItemInfo;
            int maxOccurenceRow = 0;      // счетчик строк максимального количества вхождений для элемента в заказах
            int type;                     // тип элементов номенклатуры

            if (isPKIReport) 
                type = (int)ItemType.TYPE_PKI;
            else
                type = (int)ItemType.TYPE_DETAIL;
           

            // получение номенклатуры объектов отчета
            List<ItemInfo> items = PurchaseDesignProductOperation.GetReportData(productIds, type);
            
            foreach (ItemInfo item in items)
            {
                int row1 = row;

                // вывод Id, обозначения, наименования
                xls[1, row].SetValue(item.Id);
                xls[2, row].SetValue(item.Denotation);
                xls[3, row].SetValue(item.Name);

                maxOccurenceRow = 0;

                // общее количество элемента номенклатуры в заказе
                for (int k = 0; k < productIds.Count; k++)
                {
                    int countInProductInt = PurchaseDesignProductOperation.GetObjectTotalCountInProduct(item.Id, productIds[k]);
                    string countInProductString = (countInProductInt == 0 ? string.Empty : countInProductInt.ToString());
                    xls[4 + 2 * k, row].SetValue(countInProductString);
                    xls[4 + 2 * k, row].Font.Bold = true;
                    xls[4 + 2 * k, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    xls[4 + 2 * k, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                }

                xls[4 + 2 * productIds.Count, row].SetValue(item.Amount);
                xls[4 + 2 * productIds.Count, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[4 + 2 * productIds.Count, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                // общее количество элемента номенклатуры во всех заказах
                row++;
                
                // цикл по столбцам заказов
                for (int i = 0; i < productIds.Count; i++)
                {
                    // получение списка вхождений (с количествами) для элемента номенклатуры
                    occurence = PurchaseDesignProductOperation.GetObjectOccurenceParams(item.Id, productIds[i]);

                    int occurenceRow = row;
                    
                    // вывод списка вхождений для элемента номенклатуры в отчет
                    for (int j = 0; j < occurence.Count; j++) // цикл по столбцам заказов
                    {
                        // получение параметров (наименование, обозначение и количество) элементов-родителей их вывод
                        occurenceItemInfo = PurchaseDesignProductOperation.GetObjectParams(occurence[j].ParentId, productIds[i]);
                        xls[4 + 2 * i, occurenceRow].SetValue(occurenceItemInfo.Name + " (" + occurenceItemInfo.Amount + ")");
                        xls[4 + 2 * i, occurenceRow].NoteText(occurenceItemInfo.Denotation);
                        // количество элемента номенклатуры во вхождении
                        xls[5 + 2 * i, occurenceRow].SetValue(occurence[j].Amount);

                        if (occurence.Count > 1 && j != occurence.Count - 1) occurenceRow++;
                    }                                        
                    if (maxOccurenceRow < occurenceRow) maxOccurenceRow = occurenceRow;
                }

                row = maxOccurenceRow++;
                xls[1, row1, 4 + 2 * productIds.Count, row - row1 + 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                                   XlBordersIndex.xlEdgeBottom,
                                                                                                                                   XlBordersIndex.xlEdgeLeft,
                                                                                                                                   XlBordersIndex.xlEdgeRight,
                                                                                                                                   XlBordersIndex.xlInsideVertical);
                row++;
            }
        }

        public void GetExcelReport(List<int> productIds, bool isPKIReport, string reportFilePath)
        {
            bool isCatch = false;
            xls = new Xls();

            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(reportFilePath);
                // заголовок отчета
                int row = GetExcelReportHeader(productIds, isPKIReport);
                GetExcelReportBody(productIds, isPKIReport, row);
                xls.AutoWidth();
                xls.Save();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (isCatch)
                    xls.Quit(false);
                else
                {
                    xls.Quit(true);
                    // Открытие файла
                    System.Diagnostics.Process.Start(reportFilePath);
                }
            }            
        }
    }    
}
    



