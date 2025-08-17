using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace SmallBuyAndSameNameProductReport
{
    class GetReportParameters
    {
        public static readonly Guid RKKReference = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");           //Guid Справочника РКК
        public static readonly Guid SZNOType = new Guid("2a3e0993-8cdf-4ebb-bf13-0c187dcdbc9e");               //тип СЗНО        

        public static readonly Guid confirmStageGuid = new Guid("a5ea2e1c-d441-42fd-8f92-49840351d6c1");       //стадия Утверждено
        public static readonly Guid inArchiveStageGuid = new Guid("554f7f22-9c5b-4f5f-84e4-068360678ea7");     //стадия В архиве (СЗНО)

        public static readonly Guid BuyTypeParameter = new Guid("0cda1519-5396-4d09-8c58-cf6998c68649");       //параметр Вид закупки
        //public static readonly Guid SameNameTypeGuid = new Guid("ec20eefe-1ce0-4fad-ba23-e5815f262f0b");       //параметр Вид одноименной продукции
        //public static readonly Guid SameNamePointNumGuid = new Guid("59c4cae5-caa7-412f-a39d-bc026fb2bc9a");   //параметр № пункта
        public static readonly Guid SZNOContentGuid = new Guid("d6be5cb3-b4df-40e7-9b0e-fde1c6a2ce56");        //параметр Наименование товара, услуги
        public static readonly Guid SZNODepGuid = new Guid("11e79a54-add5-4002-bdc1-2dcda453c8a6");            //параметр Номер подразделения
        public static readonly Guid SZNOCheckNumGuid = new Guid("6e527664-a3b3-4b70-b602-3421c3e05abd");       //параметр Номер счета
        public static readonly Guid SZNOCheckDateGuid = new Guid("66a80fda-0c60-47b1-a74e-d8b548e538cb");      //параметр Дата счета
        public static readonly Guid SZNOCheckSumGuid = new Guid("f962d417-de7d-4723-ba9c-ed736271ab22");       //параметр Всего сумма к оплате
        public static readonly Guid SZNOOrderNumGuid = new Guid("f430baef-49e8-4ead-a375-5edcf2f31dd7");       //параметр Номер договора
        public static readonly Guid SZNOOrderDateGuid = new Guid("5f170b49-bceb-4db1-8b28-437e316c3ebc");      //параметр Дата договора
        public static readonly Guid SZNOContractorGuid = new Guid("ed34fb5e-d8a4-4561-8f81-480edb85d11b");     //параметр Контрагент
        public static readonly Guid SZNORegNumGuid = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");         //параметр Регистрационный номер
        public static readonly Guid SZNOConfirmDateGuid = new Guid("fe8e8c1e-91f6-4e82-85ff-f12274c5aa20");    //параметр Дата утверждения
        public static readonly Guid SZNOFromGuid = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");           //параметр От

        public static readonly Guid OKDP2CodeLinkGuid = new Guid("8def4476-651f-4baa-a6aa-191e5120d2fc");      //связь Статьи ОКДП2
        public static readonly Guid OKDP2CodeGuid = new Guid("b3573c1b-412c-48f3-8c06-5f0e7aea64b0");          //параметр Код из связи Статьи



        public void FillRKK(IReportGenerationContext context, AttributeReport attributeReport)
        {
            var progressForm = new ProgressForm();
            progressForm.Visible = true;
            progressForm.progressBarMakeReport.PerformStep();
            progressForm.progressBarMakeReport.Step = 240;
            progressForm.label.Text = "Пожалуйста подождите. Формирование фильтра для СЗНО...";
            progressForm.progressBarMakeReport.PerformStep();
            //System.Windows.Forms.Application.DoEvents();

            // получение справочника для работы с данными         
            var referenceInfo = ReferenceCatalog.FindReference(RKKReference);
            var reference = referenceInfo.CreateReference();

            // получение объектов справочника          
            progressForm.progressBarMakeReport.PerformStep();
            ClassObject classObject = referenceInfo.Classes.Find(SZNOType);
            progressForm.progressBarMakeReport.PerformStep();
            System.Windows.Forms.Application.DoEvents();

            //var buyType = 1;
            //if (!attributeReport.isSmallBuyReport)
            //    buyType = 2;
            var buyType = 0;
            
            //Создаем фильтр
            var filter = new TFlex.DOCs.Model.Search.Filter(referenceInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);
            filter.Terms.AddTerm("Вид закупки", ComparisonOperator.NotEqual, buyType);
            filter.Terms.AddTerm("{"+SZNOConfirmDateGuid.ToString()+"}", ComparisonOperator.GreaterThanOrEqual, attributeReport.startDate);
            filter.Terms.AddTerm("{" + SZNOConfirmDateGuid.ToString() + "}", ComparisonOperator.LessThanOrEqual, attributeReport.endDate);
            filter.Terms.AddTerm("Стадия", ComparisonOperator.IsOneOf, new[] { Stage.Find(confirmStageGuid), Stage.Find(inArchiveStageGuid) });
            
            //Получаем список объектов по фильтру
            var listObj = reference.Find(filter);
            if (listObj.Count == 0)
            {
                MessageBox.Show("В указанном периоде не найдено ни одной СЗНО");
                progressForm.Close();
                return;
            }
            progressForm.progressBarMakeReport.PerformStep();
            //System.Windows.Forms.Application.DoEvents();
            
            Xls xls = new Xls();
            try
            {
                progressForm.label.Text = "Пожалуйста подождите. Запись данных...";

                progressForm.progressBarMakeReport.Step = 250;

                context.CopyTemplateFile();    // Создаем копию шаблона
                xls.Application.DisplayAlerts = false;

                //Формирование отчета
                MakeReport(xls, listObj, attributeReport);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                progressForm.progressBarMakeReport.PerformStep();
                System.Windows.Forms.Application.DoEvents();
                progressForm.Close();
                xls.Visible = true;
            }
        }

        public void MakeReport(Xls xls, List<ReferenceObject> pays, AttributeReport attributeReport)
        {
            int row = 1;
            int col = 1;

            if (pays.Count == 0)
                return;

            var header = "Отчет о закупках по единственному поставщику";
            //var header = "Отчет о малых закупках";
            //if (!attributeReport.isSmallBuyReport)                
            //    header = "Отчет о закупках одноименной продукции";            

            pays = pays.OrderBy(t => t[SZNOFromGuid].Value).ToList();
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 13;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row, 10, 1].MergeCells = true;
            col = 1;
            row++;

            // формирование заголовков отчета 
            /*if (attributeReport.isSmallBuyReport)
            {                
                InsertHeader("Наименование приобретаемых товаров/услуг", col, row, xls, false);
                xls[col, row, 3, 1].MergeCells = true;
                col += 3;
                InsertHeader("Отдел", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("Счет", col, row, xls, false);
                xls[col, row, 3, 1].MergeCells = true;
                col += 3;
                InsertHeader("Договор", col, row, xls, false);
                xls[col, row, 2, 1].MergeCells = true;
                col += 2;
                InsertHeader(" ", col, row, xls, false);
                xls[col, row, 2, 1].MergeCells = true;
                col += 2;
                InsertHeader("Контрагент", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("Служебная записка", col, row, xls, false);
                xls[col, row, 2, 1].MergeCells = true;
                
                xls[1, row, col+1, 1].SetBorders(
                    XlLineStyle.xlContinuous,
                    XlBorderWeight.xlThin,
                    XlBordersIndex.xlEdgeTop,
                    XlBordersIndex.xlEdgeBottom,
                    XlBordersIndex.xlEdgeLeft,
                    XlBordersIndex.xlEdgeRight,
                    XlBordersIndex.xlInsideVertical);

                col = 1; row++;
                InsertHeader("Код ОКДП2", col, row, xls, false);
                col++;
                InsertHeader("ст.36 ч.2. пункт", col, row, xls, false);
                col++;
                InsertHeader("наименование товара, услуги", col, row, xls, false);
                col += 2;
                InsertHeader("№", col, row, xls, false);
                col++;
                InsertHeader("дата", col, row, xls, false);
                col++;
                InsertHeader("сумма", col, row, xls, false);
                col++;
                InsertHeader("№", col, row, xls, false);
                col++;
                InsertHeader("дата", col, row, xls, false);
                col++;
                InsertHeader("СМСП/С(секретность)", col, row, xls, false);
                col++;
                InsertHeader("СМСП в соотв-вии с перечнем закупок у СМСП", col, row, xls, false);
                col += 2;
                InsertHeader("№", col, row, xls, false);
                col++;
                InsertHeader("дата заключения договора/оплаты счета", col, row, xls, false);

                xls[1, row, col, 1].SetBorders(
                    XlLineStyle.xlContinuous,
                    XlBorderWeight.xlThin,
                    XlBordersIndex.xlEdgeTop,
                    XlBordersIndex.xlEdgeBottom,
                    XlBordersIndex.xlEdgeLeft,
                    XlBordersIndex.xlEdgeRight,
                    XlBordersIndex.xlInsideVertical);
            }
            else*/
            {
                InsertHeader("№", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("Наименование приобретаемых товаров/услуг", col, row, xls, false);
                xls[col, row, 3, 1].MergeCells = true;
                col += 3;
                InsertHeader("Отдел", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("Счет", col, row, xls, false);
                xls[col, row, 3, 1].MergeCells = true;
                col += 3;
                InsertHeader("Договор", col, row, xls, false);
                xls[col, row, 2, 1].MergeCells = true;
                col += 2;
                InsertHeader("квартал", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("Контрагент", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("СМП", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("СМП по перечню", col, row, xls, false);
                xls[col, row, 1, 2].MergeCells = true;
                col++;
                InsertHeader("Служебная записка", col, row, xls, false);
                xls[col, row, 2, 1].MergeCells = true;

                xls[1, row, col + 1, 1].SetBorders(
                    XlLineStyle.xlContinuous,
                    XlBorderWeight.xlThin,
                    XlBordersIndex.xlEdgeTop,
                    XlBordersIndex.xlEdgeBottom,
                    XlBordersIndex.xlEdgeLeft,
                    XlBordersIndex.xlEdgeRight,
                    XlBordersIndex.xlInsideVertical);

                col = 2; row++;
                InsertHeader("Код ОКДП2", col, row, xls, false);
                col++;
                InsertHeader("ст.36 ч.2. пункт", col, row, xls, false);
                col++;
                InsertHeader("наименование товара, услуги", col, row, xls, false);
                col += 2;
                InsertHeader("№", col, row, xls, false);
                col++;
                InsertHeader("дата", col, row, xls, false);
                col++;
                InsertHeader("сумма", col, row, xls, false);
                col++;
                InsertHeader("№", col, row, xls, false);
                col++;
                InsertHeader("дата", col, row, xls, false);
                col += 3;
                InsertHeader("СМП", col, row, xls, false);
                col++;
                InsertHeader("СМП по перечню", col, row, xls, false);
                col++;
                InsertHeader("№", col, row, xls, false);
                col++;
                InsertHeader("дата заключения договора/оплаты счета", col, row, xls, false);

                xls[1, row, col, 1].SetBorders(
                    XlLineStyle.xlContinuous,
                    XlBorderWeight.xlThin,
                    XlBordersIndex.xlEdgeTop,
                    XlBordersIndex.xlEdgeBottom,
                    XlBordersIndex.xlEdgeLeft,
                    XlBordersIndex.xlEdgeRight,
                    XlBordersIndex.xlInsideVertical);
            }
            row++;
            
            int number = 0;

            /*if (attributeReport.isSmallBuyReport)
            {
                foreach (var pay in pays)
                {
                    col = 1;
                    xls[col, row].NumberFormat = "@";
                    if (pay.GetObject(OKDP2CodeLinkGuid) != null)
                    {
                        xls[col, row].SetValue(pay.GetObject(OKDP2CodeLinkGuid)[OKDP2CodeGuid].GetString());
                    }
                    else
                    {
                        xls[col, row].SetValue(" ");
                    }
                    col++;
                    xls[col, row].NumberFormat = "@";
                    var parameterBuyType = pay[BuyTypeParameter];
                    xls[col, row].SetValue(parameterBuyType.ParameterInfo.ValueList.GetName(parameterBuyType.Value));
                    col++;
                    xls[col, row].SetValue(pay[SZNOContentGuid].ToString());
                    col++;
                    xls[col, row].SetValue(pay[SZNODepGuid].ToString());
                    col++;
                    xls[col, row].NumberFormat = "000000";
                    xls[col, row].SetValue(pay[SZNOCheckNumGuid].GetString());
                    col++;
                    xls[col, row].SetValue(Convert.ToDateTime(pay[SZNOCheckDateGuid].Value).Date);
                    col++;
                    xls[col, row].NumberFormat = "#,##0.00 _?";
                    xls[col, row].SetValue(Convert.ToDouble(pay[SZNOCheckSumGuid].Value));
                    col++;
                    xls[col, row].NumberFormat = "000000";
                    xls[col, row].SetValue(pay[SZNOOrderNumGuid].GetString());
                    col++;
                    xls[col, row].SetValue(Convert.ToDateTime(pay[SZNOOrderDateGuid].Value).Date);
                    col += 3;
                    xls[col, row].SetValue(pay[SZNOContractorGuid].ToString());
                    col++;
                    xls[col, row].SetValue(pay[SZNORegNumGuid].ToString());
                    col++;
                    xls[col, row].SetValue(Convert.ToDateTime(pay[SZNOFromGuid].Value).Date);

                    xls[1, row, col, 1].SetBorders(
                    XlLineStyle.xlContinuous,
                    XlBorderWeight.xlThin,
                    XlBordersIndex.xlEdgeTop,
                    XlBordersIndex.xlEdgeBottom,
                    XlBordersIndex.xlEdgeLeft,
                    XlBordersIndex.xlEdgeRight,
                    XlBordersIndex.xlInsideVertical);
                    col++;
                    row++;
                }
            }
            else*/
            {
                foreach (var pay in pays)
                {
                    number++;
                    col = 1;
                    xls[col, row].SetValue(number++);
                    col++;
                    xls[col, row].NumberFormat = "@";
                    if (pay.GetObject(OKDP2CodeLinkGuid) != null)
                    {
                        xls[col, row].SetValue(pay.GetObject(OKDP2CodeLinkGuid)[OKDP2CodeGuid].GetString());
                    }
                    else
                    {
                        xls[col, row].SetValue(" ");
                    }
                    col++;
                    xls[col, row].NumberFormat = "@";
                    var parameterBuyType = pay[BuyTypeParameter];
                    xls[col, row].SetValue(parameterBuyType.ParameterInfo.ValueList.GetName(parameterBuyType.Value));
                    col++;
                    xls[col, row].SetValue(pay[SZNOContentGuid].ToString());
                    col++;
                    xls[col, row].SetValue(pay[SZNODepGuid].ToString());
                    col++;
                    xls[col, row].NumberFormat = "000000";
                    xls[col, row].SetValue(pay[SZNOCheckNumGuid].GetString());
                    col++;
                    xls[col, row].SetValue(Convert.ToDateTime(pay[SZNOCheckDateGuid].Value).Date);
                    col++;
                    xls[col, row].NumberFormat = "#,##0.00 _?";
                    xls[col, row].SetValue(Convert.ToDouble(pay[SZNOCheckSumGuid].Value));
                    col++;
                    xls[col, row].NumberFormat = "000000";
                    xls[col, row].SetValue(pay[SZNOOrderNumGuid].GetString());
                    col++;
                    xls[col, row].SetValue(Convert.ToDateTime(pay[SZNOOrderDateGuid].Value).Date);
                    col +=2;
                    xls[col, row].SetValue(pay[SZNOContractorGuid].ToString());
                    col +=3;
                    xls[col, row].SetValue(pay[SZNORegNumGuid].ToString());
                    col++;
                    xls[col, row].SetValue(Convert.ToDateTime(pay[SZNOFromGuid].Value).Date);

                    xls[1, row, col, 1].SetBorders(
                    XlLineStyle.xlContinuous,
                    XlBorderWeight.xlThin,
                    XlBordersIndex.xlEdgeTop,
                    XlBordersIndex.xlEdgeBottom,
                    XlBordersIndex.xlEdgeLeft,
                    XlBordersIndex.xlEdgeRight,
                    XlBordersIndex.xlInsideVertical);
                    col++;
                    row++;
                }
            }

            xls.AutoWidth();
            xls[1, 1, 1, row].EntireRow.AutoFit();
        }

        // Запись в ячейку элемента заголовка таблицы
        public static void InsertHeader(string header, int col, int row, Xls xls, bool wrap)
        {
            xls[col, row].WrapText = wrap;

            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
        }
        public static string text(string[] textArray)
        {
            string str = string.Empty;
            if (textArray.Count() > 1)
            {
                for (int i = 0; i < textArray.Count() - 1; i++)
                {
                    if (textArray[i].Trim() != string.Empty)
                        str += textArray[i] + '\n';
                }
                str += textArray[textArray.Count() - 1];
            }
            else str = textArray[textArray.Count() - 1];

            return str;
        }
    }
}
