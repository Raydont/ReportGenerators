using System;
using System.Collections.Generic;
using TFlex.Reporting;
using System.Windows.Forms;
using System.Data.SqlClient;
using ReportHelpers;
using Microsoft.Office.Interop.Excel;

namespace NomenclatureReport
{
    public class ExcelGenerator : IReportGenerator
    {
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

        //Создание экземпляра формы выбора параметров формирования отчета
        public SelectionForm selectionForm = new SelectionForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            context.CopyTemplateFile();    // Создаем копию шаблона
            selectionForm.DotsEntry(context);
        }
    }

    // ТЕКСТ МАКРОСА ===================================
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
    public class NomenclatureReport
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

            addedProducts.Columns.Clear();
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
            DataGridViewTextBoxColumn countAddedProductsColumn = new DataGridViewTextBoxColumn();
            addedProducts.Columns.Add(countAddedProductsColumn);
            countAddedProductsColumn.HeaderText = "Количество";


        }

        // заполнение DataGridView информацией о комплексах
        public void FillProducts(List<ItemInfo> complexes, DataGridView dataGridView)
        {
            dataGridView.Rows.Clear();

            foreach (ItemInfo complex in complexes)
                dataGridView.Rows.Add(complex.Id, complex.Name, complex.Denotation, complex.UnitCode, complex.ClassId, 1);

            dataGridView.AutoResizeColumns();
        }

        public void AddProduct( DataGridView dataGridView, int docId)
        {
            TFDDocument selectedComplex = new TFDDocument();

            using (SqlConnection connection = PurchaseDesignProductOperation.GetConnectionTFLEX())
            {
                connection.Open();

                string sqlRequest = string.Empty;

                // получение параметров комплекса
                string complexParamQuery = String.Format(SQLRequest.NameComplexWithCode, docId);
                SqlCommand complexParamCommand = new SqlCommand(complexParamQuery, connection);
                complexParamCommand.CommandTimeout = 0;
                SqlDataReader reader = complexParamCommand.ExecuteReader();

                while (reader.Read())
                {
                    selectedComplex.Naimenovanie = reader.GetString(0);
                    selectedComplex.Denotation = reader.GetString(1);
                    selectedComplex.Code = reader.GetString(2);
                }
                reader.Close();
            }
            dataGridView.Rows.Add( docId, selectedComplex.Code, selectedComplex.Naimenovanie, selectedComplex.Denotation,"1");

            dataGridView.AutoResizeColumns();
        }
 
        // Группировка элементов по типам с сортировкой внутри групп
        public  List<TFDDocument> SortData(List<TFDDocument> reportData, ReportParameters reportParameters)
        {
            List<TFDDocument> sortedReportData = new List<TFDDocument>();

            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Assembly, reportParameters.Assembly));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.StandartItem, reportParameters.Standart));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Detail, reportParameters.Detail));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.OtherItem, reportParameters.Other));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Material, reportParameters.Materials));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Complement, reportParameters.Complement));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Complex, reportParameters.Complex));

            return sortedReportData;
        }

        //Сортировка внутри группы объектов одного и того же типа
        public List<TFDDocument> SortOneTypeObjects(List<TFDDocument> reportData, int type, bool selectType)
        {
            List<TFDDocument> oneTypeObjects = new List<TFDDocument>();
            if (selectType)
            {
                foreach (TFDDocument doc in reportData)
                    if (doc.Class == type)
                        oneTypeObjects.Add(doc);

                oneTypeObjects.Sort(new NaimenAndOboznachComparer()); //сортируем записи по обозначению и наименованию
            }
            return oneTypeObjects;
        }

        // заголовок отчета Excel
        public int GetExcelReportHeader(PurchaseDesignProductOperation productOperation)
        {
            xls[3, 1].SetValue("Расчет общего количества изделий");
            xls[3, 1].Font.Bold = true;
            xls[3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[3, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
/*            excelApp.setValueCell(1, 3, "Расчет общего количества изделий");
            excelApp.Bold(1, 3, true);
            excelApp.centerTextCell(1, 3);*/

            xls[1, 3].SetValue("Перечень заказов");
            xls[1, 3].Font.Underline = true;
/*            excelApp.setValueCell(3, 1, "Перечень заказов");
            excelApp.Underline(3, 1, true);*/

            xls[1, 4].SetValue("Шифр");
            xls[2, 4].SetValue("Обозначение");
            xls[3, 4].SetValue("Наименование");
            xls[1, 4, 3, 1].Font.Bold = true;
            xls[1, 4, 3, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 4, 3, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, 4, 3, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);

/*            excelApp.setValueCell(4, 1, "Шифр");
            excelApp.Bold(4, 1, true); excelApp.centerTextCell(4, 1);

            excelApp.BorderCells(4, 1, 4, 1, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin);

            excelApp.setValueCell(4, 2, "Обозначение");
            excelApp.Bold(4, 2, true); excelApp.centerTextCell(4, 2);
            excelApp.BorderCells(4, 2, 4, 2, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin);

            excelApp.setValueCell(4, 3, "Наименование");
            excelApp.Bold(4, 3, true); excelApp.centerTextCell(4, 3);
            excelApp.BorderCells(4, 3, 4, 3, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin);*/

            int row = 5;

            foreach (TFDDocument doc in productOperation.listComplex)
            {
                xls[1, row].SetValue(doc.Code);
                xls[2, row].SetValue(doc.Denotation);
                xls[3, row].SetValue(doc.Naimenovanie);
/*                excelApp.setValueCell(row, 1, doc.Code);
                excelApp.setValueCell(row, 2, doc.Denotation);
                excelApp.setValueCell(row, 3, doc.Naimenovanie);*/
                row++;
            }
            xls[1, 5, 3, row - 4].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);
/*            excelApp.BorderLineCells(5, 1, row, 3,
                   Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                   Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                   Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                   Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight,
                   Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical);*/
            row += 2;

            xls[1, row].SetValue("ID");
            xls[2, row].SetValue("Обозначение");
            xls[3, row].SetValue("Наименование");
            xls[4, row].SetValue("Шифр");
            xls[1, row, 4, 1].Font.Bold = true;
            xls[1, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 4, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, row, 4, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            
/*            excelApp.setValueCell(row, 1, "ID");
            excelApp.centerTextCell(row, 1); excelApp.Bold(row, 1, true); excelApp.BorderCells(row, 1, row, 1, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick);

            excelApp.setValueCell(row, 2, "Обозначение");
            excelApp.centerTextCell(row, 2); excelApp.Bold(row, 2, true); excelApp.BorderCells(row, 2, row, 2, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick);

            excelApp.setValueCell(row, 3, "Наименование");
            excelApp.centerTextCell(row, 3); excelApp.Bold(row, 3, true); excelApp.BorderCells(row, 3, row, 3, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick);

            excelApp.setValueCell(row, 4, "Шифр");
            excelApp.centerTextCell(row, 4); excelApp.Bold(row, 4, true); excelApp.BorderCells(row, 4, row, 4, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick);*/

            // заголовки заказов
            int col = 5;
            for (int i = 0; i < productOperation.listComplex.Count; i++)
            {
                col = 5 + i;
                xls[col, row].SetValue(productOperation.listComplex[i].Code);
                xls[col, row].Columns.AutoFit();
                xls[col, row].Font.Bold = true;
                xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                xls[col, row].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
/*                excelApp.setValueCell(row, col, productOperation.listComplex[i].Code);
                excelApp.AutoWidth();
                excelApp.centerTextCell(row, col); excelApp.Bold(row, col, true); excelApp.BorderCells(row, col, row, col, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick);*/
            }

            col++;
            xls[col, row].SetValue("Общее количество");
            xls[col, row].Font.Bold = true;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
/*            excelApp.setValueCell(row, col, "Общее количество");
            excelApp.centerTextCell(row, col); excelApp.Bold(row, col, true); excelApp.BorderCells(row, col, row, col, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick);*/

            xls[1, row + 1, col, 1].Select();
            xls.Application.ActiveWindow.FreezePanes = true;
//            excelApp.FreezeArea(row + 1, 1, row + 1, col);

            return ++row;
        }

        // получение отчета Excel
        public void GetExcelReportBody(List<TFDDocument> listDocument, ReportParameters reportParameters, PurchaseDesignProductOperation productOperation, ProgressBar progressBar, int row)
        {
            List<TFDDocument> products = new List<TFDDocument>();

            products = SortData(listDocument, reportParameters);
          // получение номенклатуры объектов отчета
            int row1 = row;
            int beginRow = row+1;
            double totalCount = 0;
            bool header = false;

            xls[3, row].SetValue(TFDClass.InsertSection(products[0].Class));
            xls[3, row].Font.Bold = true;
            xls[3, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[3, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
/*            excelApp.setValueCell(row, 3, TFDClass.InsertSection(products[0].Class));
            excelApp.centerTextCell(row, 3); excelApp.Bold(row, 3, true);*/

            row++;
          
            totalCount = products[0].TotalCount;
            xls[1, row].SetValue(products[0].ObjectID);
            xls[2, row].SetValue(products[0].Denotation);
            xls[3, row].SetValue(products[0].Naimenovanie);
            xls[4, row].SetValue(products[0].Code);
/*            excelApp.setValueCell(row, 1, products[0].ObjectID);
            excelApp.setValueCell(row, 2, products[0].Denotation);
            excelApp.setValueCell(row, 3, products[0].Naimenovanie);
            excelApp.setValueCell(row, 4, products[0].Code);*/

            double step = 3 * progressBar.Maximum / (4 * products.Count);
            int step1 = Convert.ToInt32(Math.Round(step));

            for (int i = 0; i < products.Count; i++)
            {
                progressBar.Value += step1;
               
                row1 = row;
                if ((i > 0) && (products[i].Class != products[i - 1].Class))
                {
                    xls[3, row].SetValue(TFDClass.InsertSection(products[i].Class));
                    xls[3, row].Font.Bold = true;
                    xls[3, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    xls[3, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    xls[1, row1 - 1, 5 + productOperation.listComplex.Count, row - row1 + 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                                                         XlBordersIndex.xlEdgeBottom,
                                                                                                                                                         XlBordersIndex.xlEdgeLeft,
                                                                                                                                                         XlBordersIndex.xlEdgeRight,
                                                                                                                                                         XlBordersIndex.xlInsideVertical);
/*                    excelApp.setValueCell(row, 3, TFDClass.InsertSection(products[i].Class));
                    excelApp.centerTextCell(row, 3); excelApp.Bold(row, 3, true);
                    excelApp.BorderLineCells(row1 - 1, 1, row, 5 + productOperation.listComplex.Count,
                          Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                          Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                          Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                          Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight,
                          Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical);*/
                    header = true;
                    row++;
                }
                else 
                    header = false;


                if ((i > 0) && (products[i].ObjectID != products[i - 1].ObjectID))
                {
                    if (header)
                        row--;
                    xls[5 + productOperation.listComplex.Count, row - 1].SetValue(totalCount);
                    xls[5 + productOperation.listComplex.Count, row - 1].Font.Bold = true;
                    xls[5 + productOperation.listComplex.Count, row - 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    xls[5 + productOperation.listComplex.Count, row - 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
/*                    excelApp.setValueCell(row - 1, 5 + productOperation.listComplex.Count, totalCount);
                    excelApp.centerTextCell(row - 1, 5 + productOperation.listComplex.Count); excelApp.Bold(row - 1, 5 + productOperation.listComplex.Count, true);*/

                    if (header)
                        row++;
                    totalCount = products[i].TotalCount;
                    xls[1, row].SetValue(products[i].ObjectID);
                    xls[2, row].SetValue(products[i].Denotation);
                    xls[3, row].SetValue(products[i].Naimenovanie);
                    xls[4, row].SetValue(products[i].Code);
/*                    excelApp.setValueCell(row, 1, products[i].ObjectID);
                    excelApp.setValueCell(row, 2, products[i].Denotation);
                    excelApp.setValueCell(row, 3, products[i].Naimenovanie);
                    excelApp.setValueCell(row, 4, products[i].Code);*/
                }
                else
                {
                    if (i > 0)
                    {
                        row--;
                        totalCount += products[i].TotalCount;
                        for (int j = 0; j < productOperation.listComplex.Count; j++)
                        {
                            if (products[i].productId == productOperation.listComplex[j].ObjectID)
                            {
                                double countInProductInt = products[i].TotalCount;
                                string countInProductString = (countInProductInt == 0 ? string.Empty : countInProductInt.ToString());
                                xls[5 + j, row].SetValue(countInProductString);
                                xls[5 + j, row].Font.Bold = true;
                                xls[5 + j, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                xls[5 + j, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
/*                                excelApp.setValueCell(row, 5 + j, countInProductString);
                                excelApp.centerTextCell(row, 5 + j); excelApp.Bold(row, 5 + j, true);*/
                            }
                        }           
                    }
                }

                for (int j = 0; j < productOperation.listComplex.Count; j++)
                {
                    if (products[i].productId == productOperation.listComplex[j].ObjectID)
                    {
                        double countInProductInt = products[i].TotalCount;
                        string countInProductString = (countInProductInt == 0 ? string.Empty : countInProductInt.ToString());
                        xls[5 + j, row].SetValue(countInProductString);
                        xls[5 + j, row].Font.Bold = true;
                        xls[5 + j, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        xls[5 + j, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
/*                        excelApp.setValueCell(row, 5 + j, countInProductString);
                        excelApp.centerTextCell(row, 5 + j); excelApp.Bold(row, 5+j, true);*/
                    }
                }               
               
                // общее количество элемента номенклатуры в заказе
                xls[1, row1 - 1, 5 + productOperation.listComplex.Count, row - row1 + 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                                                                     XlBordersIndex.xlEdgeRight,
                                                                                                                                                     XlBordersIndex.xlInsideVertical);
/*                excelApp.BorderLineCells(row1 - 1, 1, row, 5 + productOperation.listComplex.Count,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical);*/
                row++;
            }
            xls[5 + productOperation.listComplex.Count, row - 1].SetValue(totalCount);
            xls[5 + productOperation.listComplex.Count, row - 1].Font.Bold = true;
            xls[5 + productOperation.listComplex.Count, row - 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[5 + productOperation.listComplex.Count, row - 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
/*            excelApp.setValueCell(row - 1, 5 + productOperation.listComplex.Count, totalCount);
            excelApp.centerTextCell(row - 1, 5 + productOperation.listComplex.Count); excelApp.Bold(row - 1, 5 + productOperation.listComplex.Count, true);*/
        }

        public void GetExcelReport(List<TFDDocument> listDocument, string reportFilePath, ReportParameters reportParameters, PurchaseDesignProductOperation productOperation, ProgressBar progressBar)
        {
            bool isCatch = false;
            xls = new Xls();

            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(reportFilePath);
                // заголовок отчета
                int row = GetExcelReportHeader(productOperation);
                // тело отчета
                GetExcelReportBody(listDocument, reportParameters, productOperation, progressBar, row);
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