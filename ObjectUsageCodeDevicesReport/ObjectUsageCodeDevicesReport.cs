using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.Reporting;

namespace ObjectUsageCodeDevicesReport
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

            selectionForm.Init(context);
            if (selectionForm.ShowDialog() == DialogResult.OK)
            {
                selectionForm.MakeReport();
            }
        }

        public delegate void LogDelegate(string line);

    }
    public class ObjectUsageCodeDevicesReport
    {

        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            ObjectUsageCodeDevicesReport report = new ObjectUsageCodeDevicesReport();
            report.Make(context, reportParameters);
        }

        public void Make(IReportGenerationContext context, ReportParameters reportParameters)
        {
            context.CopyTemplateFile();    // Создаем копию шаблона

            m_form.TopMost = true;
            m_form.Visible = true;
            m_form.Activate();
            ExcelGenerator.LogDelegate m_writeToLog;
            m_writeToLog = new ExcelGenerator.LogDelegate(m_form.writeToLog);
            m_form.progressBarParam(2);

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            Xls xls = new Xls();
            bool isCatch = false;

            m_form.setStage(IndicationForm.Stages.DataAcquisition);
            m_writeToLog("=== Получение данных ===");

            MakerReport(reportParameters, m_writeToLog);

            m_form.Dispose();
        }


        public void MakerReport(ReportParameters reportParameters, ExcelGenerator.LogDelegate m_writeToLog)
        {
            List<List<string>> branches = new List<List<string>>();

            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;

            int row = 1;
            //   MessageBox.Show(reportParameters.PKIObjects.Count+"");
            //   errorId = 1;
            var countSepForObj = Convert.ToInt32(Math.Round((double)(100 / reportParameters.PKIObjects.Count), 0));


            foreach (var obj in reportParameters.PKIObjects)
            {
                branches = GetParent(obj, new List<string>()).Distinct().ToList();
                //if (branches == null)
                //    break;
                m_writeToLog("Запись для " + obj[NameNomenclature].Value.ToString());

                if (reportParameters.FullInformation)
                {
                    row = WriteObjects(xls, branches, obj[NameNomenclature].Value.ToString(), row, countSepForObj);
                }
                else
                {
                    row = WriteObjectsHighLevel(xls, branches, obj[NameNomenclature].Value.ToString(), row, countSepForObj);
                }
            }

            xls.AutoWidth();
            xls.AutoHeight();
            xls.Visible = true;
        }

        public readonly Guid DenotationNomenclature = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
        public readonly Guid NameNomenclature = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
        public List<List<string>> GetParent(ReferenceObject refObject, List<string> mainBranch)
        {
            List<List<string>> branches = new List<List<string>>();

            foreach (var parent in refObject.Parents.ToList())
            {
                // Если родитель = Сборрочная единица
                if (parent.Class.IsInherit(new NomenclatureReference().Classes.AllClasses.Find("Сборочная единица").Guid)
                    || parent.Class.IsInherit(new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Guid))
                {
                    //Если родитель шифрованное устройство, то записываем ветку
                    if (parent[new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6")].Value.ToString().Trim() != string.Empty
                        || parent.Class.IsInherit(new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Guid))
                    {
                        var branch = new List<string>();
                        branch.AddRange(mainBranch);
                        branch.Add(parent[NameNomenclature].Value + " " + parent[DenotationNomenclature].Value);
                        branches.Add(branch);
                    }
                    else
                    {
                        var branch = new List<string>();
                        branch.AddRange(mainBranch);
                        branch.Add(parent[NameNomenclature].Value + " " + parent[DenotationNomenclature].Value);
                        branches.AddRange(GetParent(parent, branch));
                    }
                }
            }
            return branches;
        }

        public static int InsertHeader(string header, int col, int row, Xls xls)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row].Interior.Color = Color.LightGray;
            col++;

            return col;
        }

        private int WriteObjects(Xls xls, List<List<string>> branches, string pki, int row, int countStep)
        {

            InsertHeader(pki, 1, row, xls);
            var countStepBranch = Convert.ToInt32(Math.Round((double)(countStep/branches.Count)));

            var col = 1;
            row++;
            var maxLevel = branches.Max(t=>t.Count);


            //колонки обязательные для вывода
            for (int i = 1; i <= maxLevel; i++)
            {
                col = InsertHeader(i + " уровень", col, row, xls);
            }


            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);
            row++;
            foreach (var branch in branches)
            {
                m_form.progressBarParam(countStepBranch);
                col = 1;
                foreach (var obj in branch)
                {
                    SetParameter(xls, col, row, obj);
                    xls[1, row, maxLevel, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                             XlBordersIndex.xlEdgeRight,
                                                                                             XlBordersIndex.xlInsideVertical);
                    col++;
                }
                row++;

            }

            row++;
            return row;
        }

        private int WriteObjectsHighLevel(Xls xls, List<List<string>> branches, string pki, int row,  int countStep)
        {
            var countStepBranch = 1;

            if (branches.Count != 0)
            {
                countStepBranch = Convert.ToInt32(Math.Round((double)(countStep / branches.Count)));
            }
           
            InsertHeader(pki, 1, row, xls);

            var col = 1;
            //row++;

            row++;
            List<string> modifyBranches = new List<string>();

            foreach (var branch in branches)
            {
                modifyBranches.Add(branch.LastOrDefault());
            }

            foreach (var branch in modifyBranches.OrderBy(t => t).Distinct())
            {
                col = 1;
                m_form.progressBarParam(countStepBranch);
                SetParameter(xls, col, row, branch);
                xls[1, row, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                         XlBordersIndex.xlEdgeBottom,
                                                                                         XlBordersIndex.xlEdgeLeft,
                                                                                         XlBordersIndex.xlEdgeRight,
                                                                                         XlBordersIndex.xlInsideVertical);
                col++;
                row++;

            }
            row++;
            return row;
        }

        private static void SetParameter(Xls xls, int col, int row, string value)
        {
            xls[col, row].NumberFormat = "@";
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[col, row].SetValue(value);
        }

        private static void SetParameter(Xls xls, int col, int row, double value)
        {
            xls[col, row].NumberFormat = "@";
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[col, row].SetValue(value);
        }
    }
}
