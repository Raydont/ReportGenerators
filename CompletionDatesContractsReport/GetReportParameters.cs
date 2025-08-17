using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Classes;
using TFlex.Reporting;
using Microsoft.Office.Interop.Excel;
using Filter = TFlex.DOCs.Model.Search.Filter;
using Globus.DOCs.Xlsx;

namespace CompletionDatesContractsReport
{
    class GetReportParameters
    {
        public static readonly Guid RKKReference = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");  //Guid Справочника РКК
        public static readonly Guid RKKAbstractReferenceGuid = new Guid("ec785091-5483-4d5a-9cf5-188b15cd6cac");  //Guid абстрактного типа РКК 
        public static readonly Guid DZUType = new Guid("1614e166-81d1-492a-942e-4b03970fcf8a");   // тип Договор закупки услуги

        // Guids для ДЗУ
        public static readonly Guid RegNumberGuid = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");      // Guid параметра "Регистрационный номер" справочника "Регистрационно-контрольные карточки"
        public static readonly Guid CountractorGuid = new Guid("a2d8a1f1-0e4f-41f0-ad7d-512fe751d157"); //Guid Контрагента
        public static readonly Guid ContractNumberGuid = new Guid("40164d61-5172-4e85-9ab8-51ac1d27f1e4"); //Номер Договора
        public static readonly Guid FromDZUGuid = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");           // Guid параметра "От" справочника "Регистрационно-контрольные карточки"
        public static readonly Guid DateDZUGuid = new Guid("89a50a7f-2db1-48d7-9b84-e24afa2a5da7");        //Дата договора
        public static readonly Guid DepartmentGuid = new Guid("f681d30c-fd0a-4f6d-820e-4676adb33b21");    //Номер подразделения
        public static readonly Guid SummaryGuid = new Guid("bc62cf57-65a5-47e2-a3e2-7dc38c47c25a");       //Краткое содержание
        public static readonly Guid ListLinkDocumentGuid = new Guid("aab1cd65-fe2d-4fc0-8db2-4ff2b3d215d5"); //Связь ДЗУ с документами

        public static readonly Guid SummaContractGuid = new Guid("5f357c2c-58bb-456e-86bd-38b1abd7e42c");  //Сумма с НДС
        public static readonly Guid ContractClosed = new Guid("3d198691-6313-4d2d-a625-17fdf11fd335");     //Договор закрыт
        public static readonly Guid DurationAgreement = new Guid("2311154f-c7d9-436e-9c92-6d8fc9763958");  //Срок действия по договору

        public static readonly Guid ResponsibleGuid = new Guid("ec9c400c-a972-453e-b0c8-3948ba981b17");  //Ответственный

        public static readonly Guid NumberDepartment = new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420"); //Номер подразделения типв Производственное подразделение
        ProgressForm progressForm = new ProgressForm();
        public void FillRKK(IReportGenerationContext context, AttributeReport attributeReport)
        {
            System.Windows.Forms.Application.DoEvents();
            progressForm.Visible = true;
            // progressForm.progressBarMakeReport.PerformStep();
            progressForm.progressBarMakeReport.Step = 1;
            progressForm.label.Text = "Пожалуйста подождите. Формирование фильтра для ДЗУ...";
            progressForm.progressBarMakeReport.PerformStep();

            // получение справочника для работы с данными         
            var referenceInfo = ReferenceCatalog.FindReference(RKKReference);
            var reference = referenceInfo.CreateReference();

            // получение объектов справочника          
            progressForm.progressBarMakeReport.PerformStep();
            ClassObject classObject = referenceInfo.Classes.Find(DZUType);
            XlsxDocument xls = new XlsxDocument();
            progressForm.progressBarMakeReport.PerformStep();

            reference.LoadSettings.AddAllParameters();
            reference.LoadSettings.AddRelation(ListLinkDocumentGuid);


            var sumValue = 100000;
            if (attributeReport.SumOrder500000)
                sumValue = 500000;
            System.Windows.Forms.Application.DoEvents();
            //Создаем фильтр
            var filterDZU2Weeks = new Filter(referenceInfo);

            var filterDZUDelayed = new Filter(referenceInfo);

            filterDZUDelayed.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);
            filterDZU2Weeks.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);


            filterDZUDelayed.Terms.AddTerm("Договор закрыт", ComparisonOperator.Equal, false);
            filterDZUDelayed.Terms.AddTerm("Сумма с НДС (ДЗУ)", ComparisonOperator.GreaterThanOrEqual, sumValue);
            filterDZUDelayed.Terms.AddTerm("[" + DurationAgreement + "]", ComparisonOperator.LessThanOrEqual, DateTime.Now.Date);

            filterDZU2Weeks.Terms.AddTerm("Договор закрыт", ComparisonOperator.Equal, false);
            filterDZU2Weeks.Terms.AddTerm("Сумма с НДС (ДЗУ)", ComparisonOperator.GreaterThanOrEqual, sumValue);
            filterDZU2Weeks.Terms.AddTerm("[" + DurationAgreement + "]", ComparisonOperator.GreaterThan, DateTime.Now.Date);
            filterDZU2Weeks.Terms.AddTerm("[" + DurationAgreement + "]", ComparisonOperator.LessThanOrEqual, DateTime.Now.Date.AddDays(14));

            if (attributeReport.DepartmentReport && attributeReport.Department != null)
            {
                filterDZUDelayed.Terms.AddTerm("[" + DepartmentGuid + "]", ComparisonOperator.Equal, attributeReport.Department[NumberDepartment].Value.ToString());
                filterDZU2Weeks.Terms.AddTerm("[" + DepartmentGuid + "]", ComparisonOperator.Equal, attributeReport.Department[NumberDepartment].Value.ToString());
            }

            if (attributeReport.MakerReport && attributeReport.Maker != null)
            {
                filterDZUDelayed.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Author], ComparisonOperator.Equal, attributeReport.Maker);
                filterDZU2Weeks.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Author], ComparisonOperator.Equal, attributeReport.Maker);
            }


            var listDZU2Week = reference.Find(filterDZU2Weeks);
            var listDZUDelayed = reference.Find(filterDZUDelayed);
            reference.LoadLinks(listDZUDelayed.Union(listDZU2Week).ToList(), reference.LoadSettings);
            System.Windows.Forms.Application.DoEvents();

            progressForm.progressBarMakeReport.PerformStep();

            try
            {
                System.Windows.Forms.Application.DoEvents();
                progressForm.label.Text = "Пожалуйста подождите. Запись данных...";

                context.CopyTemplateFile();    // Создаем копию шаблона
                                               //  xls.Application.DisplayAlerts = false;

                //  xls.Open(context.ReportFilePath);

                //Формирование отчета
                MakeReport(xls, listDZU2Week, listDZUDelayed);

                //   xls.AutoWidth();
                // xls.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                progressForm.progressBarMakeReport.PerformStep();
                progressForm.Close();
                xls.SaveTo(context.ReportFilePath);
                // xls.Quit(true);
                // Открытие файла
                System.Diagnostics.Process.Start(context.ReportFilePath);
            }
        }

        public void MakeReport(XlsxDocument xls, List<ReferenceObject> weeks2contracts, List<ReferenceObject> delayedContracts)
        {
            int row = 1;
            int col = 1;

            //Задание начальных условий для progressBar
            var progressStep = Convert.ToDouble(progressForm.progressBarMakeReport.Maximum) / (weeks2contracts.Count + delayedContracts.Count);

            if (progressForm.progressBarMakeReport.Maximum < (weeks2contracts.Count + delayedContracts.Count))
            {
                progressForm.progressBarMakeReport.Step = 1;
            }
            else
            {
                progressForm.progressBarMakeReport.Step = Convert.ToInt32(Math.Round(progressStep));
            }


            double step = 0;

            row = MakeTable(xls, weeks2contracts.OrderBy(t => Convert.ToDateTime(t[DateDZUGuid].Value).Date).ToList(), col, row, "ДЗУ со сроком окончания действия меньше 2 недель:");
            MakeTable(xls, delayedContracts.OrderBy(t => Convert.ToDateTime(t[DateDZUGuid].Value).Date).ToList(), 1, row++, "Просроченные ДЗУ");
        }

        public int MakeTable(XlsxDocument xls, List<ReferenceObject> contracts, int col, int row, string header)
        {
            if (contracts.Count == 0)
                return row;
            contracts = contracts.OrderBy(t => t[DurationAgreement].Value).ToList();
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 13;
            xls[col, row].HorizontalAlignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
            xls[col, row].VerticalAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row, 10, 1].Merge();
            col = 1;
            row++;
            // формирование заголовков отчета 
            col = InsertHeader("№ п/п", col, row, xls, false);
            col = InsertHeader("№ ДЗУ", col, row, xls, false);
            col = InsertHeader("Дата ДЗУ", col, row, xls, false);
            var colContractor = col;
            col = InsertHeader("Контрагент", col, row, xls, false);
            col = InsertHeader("Подразделение", col, row, xls, true);
            var colSummary = col;
            col = InsertHeader("Краткое содержание", col, row, xls, true);
            var colNumber = col;
            col = InsertHeader("№ Договора", col, row, xls, true);
            col = InsertHeader("Сумма с НДС", col, row, xls, false);
            var colMaker = col;

            col = InsertHeader("ФИО исполнителя", col, row, xls, true);
            xls[col - 1, row].WrapText = true;
            xls[col - 1, row].ColumnWidth = 20;

            col = InsertHeader("Ответственный", col, row, xls, true);
            xls[col - 1, row].WrapText = true;
            xls[col - 1, row].ColumnWidth = 20;

            var colDoc = col;
            col = InsertHeader("Присоединенные документы", col, row, xls, false);
            col = InsertHeader("Срок действия по договору", col, row, xls, true);
            col = InsertHeader("Договор закрыт", col, row, xls, true);


            xls[1, row, col - 1, 1].SetBorders(
                 DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin,
                 XlsxBorder.Top,
                 XlsxBorder.Bottom,
                 XlsxBorder.Left,
                 XlsxBorder.Right,
                 XlsxBorder.InsideVertical);
            row++;
            int number = 0;
            var beginRow = row;
            foreach (var contract in contracts)
            {
                progressForm.progressBarMakeReport.PerformStep();
                number++;
                var regNumber = contract[RegNumberGuid].Value.ToString();
                progressForm.label.Text = "Запись объекта " + regNumber;
                System.Windows.Forms.Application.DoEvents();
                col = 1;
                xls[col, row].SetValue(number);
                col++;
                xls[col, row].SetValue(regNumber);
                col++;
                xls[col, row].SetValue(Convert.ToDateTime(contract[DateDZUGuid].Value).Date);
                col++;
                xls[col, row].SetValue(contract[CountractorGuid].Value.ToString());
                col++;
                xls[col, row].SetValue(contract[DepartmentGuid].Value.ToString());
                col++;
                xls[col, row].SetValue(contract[SummaryGuid].Value.ToString());
                col++;
                xls[col, row].SetValue(contract[ContractNumberGuid].Value.ToString());
                col++;
                xls[col, row].SetValue(Convert.ToDouble(contract[SummaContractGuid].Value));
                col++;
                xls[col, row].SetValue(contract.SystemFields.Author.ToString());
                col++;
                xls[col, row].SetValue(contract[ResponsibleGuid].Value.ToString());
                col++;
                int i = 0;

                List<string> files = new List<string>();

                foreach (var file in contract.GetObjects(ListLinkDocumentGuid).OfType<FileObject>())
                {
                    i++;
                    try
                    {
                        var fileString = file.ToString();
                        fileString = fileString.Replace(regNumber.ToString().Replace("/", "-") + " " + Convert.ToDateTime(contract[FromDZUGuid].Value).Date.ToShortDateString(), "");
                        files.Add(i.ToString() + "." + fileString.Remove(fileString.Length - 4));
                    }
                    catch
                    {

                    }
                }

                xls[col, row].SetValue(String.Join("\n", files));
                col++;

                xls[col, row].SetValue(Convert.ToDateTime(contract[DurationAgreement].Value).ToShortDateString());
                col++;
                if (!Convert.ToBoolean(contract[ContractClosed].Value))
                    xls[col, row].SetValue("открыт");
                else
                    xls[col, row].SetValue("закрыт");


                xls[1, row, col, 1].SetBorders(
                 DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin,
                 XlsxBorder.Top,
                 XlsxBorder.Bottom,
                 XlsxBorder.Left,
                 XlsxBorder.Right,
                 XlsxBorder.InsideVertical);

                row++;

            }



            xls[colContractor, row].ColumnWidth = 30;
            xls[colContractor, beginRow, 1, row - beginRow].WrapText = true;
            xls[colSummary, row].ColumnWidth = 35;
            xls[colSummary, beginRow, 1, row - beginRow].WrapText = true;
            xls[colMaker, row].ColumnWidth = 20;
            xls[colMaker, beginRow, 1, row - beginRow].WrapText = true;
            xls[colMaker+1, row].ColumnWidth = 20;
            xls[colMaker+1, beginRow, 1, row - beginRow].WrapText = true;
            xls[colNumber, row].ColumnWidth = 15;
            xls[colDoc, row].ColumnWidth = 28;
            row++;
            return row;
        }


        // Запись в ячейку элемента заголовка таблицы
        public static int InsertHeader(string header, int col, int row, XlsxDocument xls, bool wrap)
        {
            xls[col, row].WrapText = wrap;

            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].HorizontalAlignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            col++;

            return col;
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
