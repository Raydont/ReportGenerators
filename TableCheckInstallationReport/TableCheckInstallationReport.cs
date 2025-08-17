using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using TFlex.Reporting;
using TFlex;
using TFlex.Drawing;
using TFlex.Model.Model2D;
using TFlex.Model;
using Application = TFlex.Application;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Globus.DOCs.Technology.Reports;
using Globus.DOCs.Technology.Reports.CAD;
using TFlex.DOCs.Model.References.Files;
using System.Threading;

namespace TableCheckInstallationReport
{
    public class CadGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get { return ".grb"; }
        }

        public delegate void LogDelegate(string line);

        /// Определяет редактор параметров отчета ("Параметры шаблона" в контекстном меню Отчета)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="context"></param>
        /// <param name="owner"></param>
        /// <returns></returns>
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context,
                                           System.Windows.Forms.IWin32Window owner)
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
            try
            {

                selectionForm.Init(context);

                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    TableCheckInstallationReport.MakeReport(context, selectionForm.reportParameters);
                }

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK,
                                                     System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {

            }
        }

        public class TableCheckInstallationReport : IDisposable
        {
            public IndicationForm m_form = new IndicationForm();

            public void Dispose()
            {
            }

            public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
            {
                TableCheckInstallationReport report = new TableCheckInstallationReport();
                report.Make(context, reportParameters);
            }

            private IReportGenerationContext _context;

            public void Make(IReportGenerationContext context, ReportParameters reportParameters)
            {
                _context = context;
                context.CopyTemplateFile();    // Создаем копию шаблона

                m_form.Visible = true;
                LogDelegate m_writeToLog;
                m_writeToLog = new LogDelegate(m_form.writeToLog);

                m_form.setStage(IndicationForm.Stages.Initialization);
                m_writeToLog("=== Инициализация ===");

                List<TableInstallation> reportData = new List<TableInstallation>();

                try
                {
                    using (TableCheckInstallationReport report = new TableCheckInstallationReport())
                    {
                        m_form.setStage(IndicationForm.Stages.DataAcquisition);
                        List<TableInstallation> tableInstallations = new List<TableInstallation>();
                        //Чтение данных для отчета  
                        m_form.setStage(IndicationForm.Stages.DataProcessing);
                        m_writeToLog("=== Обработка данных ===");
                        string patternPosDigit = @"\d{1,}";
                        string patternLittera = @"\D{1,}";
                        Regex regexPosDigit = new Regex(patternPosDigit);
                        Regex regexLittera = new Regex(patternLittera);
                      //  Clipboard.SetText(string.Join("\r\n",report.ProcessingData(reportParameters, m_writeToLog).Select(t=>t.NumberChain)));

                        if (reportParameters.FileDataNet.Count == 0)
                        {
                            tableInstallations = report.ProcessingData(reportParameters, m_writeToLog)
                                .OrderBy(t => regexLittera.Match(t.NumberChain).Value)
                                .ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.NumberChain).Value + "0"))
                                .ToList();
                        }
                        else
                        {
                            tableInstallations = report.ProcessingDataNet(reportParameters, m_writeToLog)
                                .OrderBy(t => regexLittera.Match(t.NumberChain).Value)
                                .ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.NumberChain).Value + "0"))
                                .ToList();
                        }

                        if (tableInstallations == null)
                        {
                            m_form.Dispose();
                            return;
                        }

                        m_form.setStage(IndicationForm.Stages.ReportGenerating);
                        m_writeToLog("=== Формирование отчета ===");
                        //Формирование отчета
                        MakeSelectionByTypeReport(tableInstallations, reportParameters, m_writeToLog);

                        m_form.progressBarParam(2);

                        m_form.setStage(IndicationForm.Stages.Done);
                        m_writeToLog("=== Завершение работы ===");
                        System.Diagnostics.Process.Start(context.ReportFilePath);

                    }
                }
                catch (Exception e)
                {
                    string message = String.Format("ОШИБКА: {0}", e.ToString());
                    System.Windows.Forms.MessageBox.Show(message, "Ошибка",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
                m_form.Dispose();
            }


            // Чтение данных для отчета
            public List<TableInstallation> ProcessingData (ReportParameters reportParameters, LogDelegate logDelegate)
            {                                                     
                string denotationChainPattern = @"(?<=^317|^327)\S{1,12}(\s{3})";                                                //Обозначение цепи
             // string posDenotationPattern = @"(?<=\s{7})((([A-Z]{1,2}\d{1,3}(.\d{1,3})?)|(\d{1,4}))\s{1,}-(([A-Z]{1,})?\d{1,}|([A-Z]{1}\d{1,})|([A-Z]{1,}))|VIA|\s{6}-\d{1})";  //Поз. Обознач.
             // string posDenotationPattern = @"(?<=\s{3})((((\d.\d{1,2})|((\d?)[A-Z]{1,2}\d{1,3}(.\d{1,3})?)|(\d{1,4}))\s{1,}|(([A-Z]{1,2}\d{1,3}(.\d{1,3})?)|(\d{1,4})))-(([A-Z]{1,})?\d{1,}|([A-Z]{1}\d{1,})|([A-Z]{1,}))|VIA|\s{6}-\d{1})";  //Поз. Обознач.
                string posDenotationPattern = @"(?<=((^317|^327)\S{1,12}(\s{3})))(\s{0,}((\S{0,}\s{0,}-\S{0,})|VIA))";
                string xPattern = @"(?<=X(\s|\+\s?))\d{5,6}(?=Y)";                                                               // X
                string yPattern = @"(?<=Y(\s|\+\s?))\d{5,6}(?=X)";                                                               // Y
                string viaPattern = @"(?<=\s{6})(VIA|\s{6}-\d{1})";  //Поз. Обознач.
                string netPattern = @"(?<=^317|^327)(\S{1,12}(\s{4}))";

                var listDenotationChain = new List<string>();
                var listPosDenotation = new List<string>();
                var listX = new List<string>();
                var listY = new List<string>();
                StringBuilder strB = new StringBuilder();

                foreach (var str in reportParameters.FileData)
                {

                    var netChain = Regex.Match(str, netPattern);
                    if (netChain.Success && !netChain.Value.Contains("NET"))
                    {
                        var newStr = str.Remove(3, netChain.Value.Length);
                        newStr = newStr.Insert(3, "NET0" + netChain.Value.Trim());
                        strB.AppendLine(newStr);
                    }
                    else
                    {
                        strB.AppendLine(str);
                    }
                    if ( Regex.Match(str, denotationChainPattern).Success && !Regex.Match(str, viaPattern).Success)
                    {
                        listDenotationChain.Add(Regex.Match(str, denotationChainPattern).Value.Trim());

                        var posDenotation = Regex.Match(str, posDenotationPattern).Value.Trim();

                        if (!reportParameters.MPP && posDenotation.Contains("(") && !posDenotation.Contains(")"))
                        {
                            posDenotation = posDenotation.Insert(posDenotation.IndexOf("-"), ")");
                        }

                        listPosDenotation.Add(posDenotation);

                        listX.Add(Regex.Match(str, xPattern).Value);

                        listY.Add(Regex.Match(str, yPattern).Value);

                        if (!Regex.Match(str, posDenotationPattern).Success || !Regex.Match(str, xPattern).Success || !Regex.Match(str, yPattern).Success)
                        {
                            MessageBox.Show("Обратитесь в бюро 911. Таблица проверки монтажа не может быть сформирована!\n"
                                + "Некорректное значение в цепи " + Regex.Match(str, denotationChainPattern).Value,
                                "Ошибка в исходном файле!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return null;
                        }
                    }


                }

                using (FileStream fs = File.OpenWrite(reportParameters.newFileName))
                {
                    Byte[] info =
                        new UTF8Encoding(true).GetBytes(strB.ToString());

                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                }

                List<TableInstallation> table = new List<TableInstallation>();
                TableInstallation chainTable = new TableInstallation();

                string patternPosDigit = @"\d{1,}(?=:|\(|\.)";
                string patternLittera = @"^([A-Z]{1,}|[А-Я]{1,}|\d{1,}[A-ZА-Я]{1,})";
                string patternPosDigit2 = @"\d{1,}(?=\)|\z)";
                string patternLittera2 = @"(?<=(\(|:))[A-Z]{1,}|[А-Я]{1,}";
                string patternPosDigit3 = @"\d{1,}(?=\z)";
                string patternLittera3 = @"(?<=:)[A-Z]{1,}|[А-Я]{1,}";

                Regex regexPosDigit = new Regex(patternPosDigit);
                Regex regexLittera = new Regex(patternLittera);
                Regex regexPosDigit2 = new Regex(patternPosDigit2);
                Regex regexLittera2 = new Regex(patternLittera2);
                Regex regexPosDigit3 = new Regex(patternPosDigit3);
                Regex regexLittera3 = new Regex(patternLittera3);

                chainTable.NumberChain = listDenotationChain[0];

                for(int i = 0; i < listDenotationChain.Count; i++)
                {
                    Chain chain = new Chain();
                    chain.Add(Convert.ToInt32(listX[i]), Convert.ToInt32(listY[i]), listPosDenotation[i],reportParameters);
                    chainTable.ChainData.Add(chain);

                    if (i < listDenotationChain.Count - 1)
                    {      
                        if (listDenotationChain[i] != listDenotationChain[i+1])
                        {
                            try
                            {
                                var resultSortData = chainTable.ChainData
                                    .OrderBy(t => regexLittera.Match(t.PosDenotation).Value)
                                    .ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.PosDenotation).Value))
                                    .ThenBy(t => regexLittera2.Match(t.PosDenotation).Value)
                                    .ThenBy(t => Convert.ToInt32(regexPosDigit2.Match(t.PosDenotation).Value))
                                    .ThenBy(t => regexLittera3.Match(t.PosDenotation).Value)
                                    .ThenBy(t => Convert.ToInt32(regexPosDigit3.Match(t.PosDenotation).Value))
                                    .ToList();

                                chainTable.ChainData = new List<Chain>();
                                chainTable.ChainData.AddRange(resultSortData);
                            }
                            catch
                            {
                                try
                                {
                                    var resultSortData = chainTable.ChainData
                                        .OrderBy(t => regexLittera.Match(t.PosDenotation).Value)
                                        .ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.PosDenotation).Value)).ToList();
                                    chainTable.ChainData = new List<Chain>();
                                    chainTable.ChainData.AddRange(resultSortData);
                                }
                                catch
                                {

                                }

                            }
                            table.Add(chainTable); 
                            chainTable = new TableInstallation();
                            chainTable.NumberChain = listDenotationChain[i+1];
                        }

                    }
                }
                table.Add(chainTable);

                return table;
            }

            public List<TableInstallation> ProcessingDataNet(ReportParameters reportParameters, LogDelegate logDelegate)
            {
                string patternPosDigit = @"\d{1,}(?=:|\(|\.)";
                string patternLittera = @"^([A-Z]{1,}|[А-Я]{1,}|\d{1,}[A-ZА-Я]{1,})";
                string patternPosDigit2 = @"\d{1,}(?=\)|\z)";
                string patternLittera2 = @"(?<=(\(|:))[A-Z]{1,}|[А-Я]{1,}";
                string patternPosDigit3 = @"\d{1,}(?=\z)";
                string patternLittera3 = @"(?<=:)[A-Z]{1,}|[А-Я]{1,}";

                Regex regexPosDigit = new Regex(patternPosDigit);
                Regex regexLittera = new Regex(patternLittera);
                Regex regexPosDigit2 = new Regex(patternPosDigit2);
                Regex regexLittera2 = new Regex(patternLittera2);
                Regex regexPosDigit3 = new Regex(patternPosDigit3);
                Regex regexLittera3 = new Regex(patternLittera3);
                List<TableInstallation> table = new List<TableInstallation>();
                
                try
                {
                    foreach (string str in reportParameters.FileDataNet)
                    {
                        TableInstallation chainTable = new TableInstallation();
                        chainTable.NumberChain = str.Remove(str.IndexOf(" "));
                        List<string> chains = str.Remove(0, str.IndexOf(" ")).Trim().Split(',').ToList();
                       
                        foreach (var ch in chains.Select(t=>t.Trim()))
                        {
                            Chain chain = new Chain();
                            string posDenotation = ch.Remove(ch.IndexOf(" ")).Trim().Replace("-", ":");

                            string xy = ch.Remove(0, ch.IndexOf(" ")).Replace("(", "").Replace(")", "").Trim();
               
                            var stepX = xy.Remove(xy.IndexOf(':'));
                            stepX = stepX.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                            var stepY = xy.Remove(0, xy.IndexOf(':')+1);
                            stepY = stepY.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                            chain.Add(Convert.ToDouble(stepX), Convert.ToDouble(stepY), posDenotation, reportParameters);
                            chainTable.ChainData.Add(chain);
                        }
                        try
                        {
                            var resultSortData = chainTable.ChainData
                                .OrderBy(t => regexLittera.Match(t.PosDenotation).Value)
                                .ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.PosDenotation).Value))
                                .ThenBy(t => regexLittera2.Match(t.PosDenotation).Value)
                                .ThenBy(t => Convert.ToInt32(regexPosDigit2.Match(t.PosDenotation).Value))
                                .ThenBy(t => regexLittera3.Match(t.PosDenotation).Value)
                                .ThenBy(t => Convert.ToInt32(regexPosDigit3.Match(t.PosDenotation).Value))
                                .ToList();

                            chainTable.ChainData = new List<Chain>();
                            chainTable.ChainData.AddRange(resultSortData);
                        }
                        catch
                        {
                            try
                            {
                                var resultSortData = chainTable.ChainData
                                    .OrderBy(t => regexLittera.Match(t.PosDenotation).Value)
                                    .ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.PosDenotation).Value)).ToList();
                                chainTable.ChainData = new List<Chain>();
                                chainTable.ChainData.AddRange(resultSortData);
                            }
                            catch
                            {

                            }

                        }
                        table.Add(chainTable);
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Обратитесь в отдел 911", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                return table;
            }

            public void MakeSelectionByTypeReport(List<TableInstallation> reportData, ReportParameters reportParameters, LogDelegate logDelegate)
            {
                var applicationSessionSetup = new ApplicationSessionSetup
                {
                    Enable3D = true,
                    ReadOnly = false,
                    PromptToSaveModifiedDocuments = false,
                    EnableMacros = true,
                    EnableDOCs = true,
                    DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version13,
                    ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
                };

                if (!Application.InitSession(applicationSessionSetup))
                {
                    return;
                }
                Application.FileLinksAutoRefresh = Application.FileLinksRefreshMode.DoNotRefresh;

                var doc = Application.OpenDocument(_context.ReportFilePath);
                doc.BeginChanges("Заполнение данными");
                var reportText = doc.GetTextByName("Table");
                reportText.BeginEdit();
                var contentTable = reportText.GetFirstTable();

                logDelegate("Количество элементов отчета: " + reportData.Count);

               // Clipboard.SetText(string.Join("\r\n", reportData.Select(t=> t.NumberChain + " - " + string.Join("\r\n",  t.ChainData.Select(c=>c.PosDenotation + "("+ c.StepX +";" + c.StepY+")")))));


                if (contentTable.HasValue)
                {
                    var cadTable = new CadTable(contentTable.Value, reportText, 41, 41);
                    int countStr = 0;
                   
                    if (reportData.Count > 0)
                    {
                        foreach(var chain in reportData)
                        {                           
                            int rowCount = (int)Math.Ceiling(chain.ChainData.ToList().Count / 4f);
                           
                            for (int j = 0; j < rowCount; j++)
                            {                      
                                var row = cadTable.CreateRow();

                                if (j == 0 || (countStr%41) == 0)
                                    row.AddText(chain.NumberChain);

                                else
                                    row.AddText("");

                                for (int i = 0; i < 4; i++)
                                {
                                    if (chain.ChainData.ToList().Count > 4 * j + i)
                                    {
                                        row.AddText(chain.ChainData.ToList()[4 * j + i].PosDenotation);
                                        row.AddText(chain.ChainData.ToList()[4 * j + i].StepX.ToString());
                                        row.AddText(chain.ChainData.ToList()[4 * j + i].StepY.ToString());
                                    }

                                }
                                countStr++;
                            }
                        }             
                    }
              

                    cadTable.Apply();
                }
                reportText.EndEdit();


                // вставка листа регистрации изменений
                var changelogPage = new Page(doc);
                changelogPage.Rectangle = new Rectangle(0, 0, 210, 297);
                // временная страница для удаления фрагментов
                var tempPage = new Page(doc);
                foreach (Fragment fragment in changelogPage.GetFragments())
                {
                    fragment.Page = tempPage;
                }
                doc.DeletePage(tempPage, new DeleteOptions(true) { DeletePageObjects = true });

                var fileReference = new FileReference();

                var fileObject = fileReference.FindByRelativePath(
                    @"Служебные файлы\Шаблоны отчётов\Таблица проверки монтажа\Лист регистрации изменений ГОСТ 2.503.grb");
                fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
                var changelogFragment = new Fragment(new FileLink(doc, fileObject.LocalPath, true));
                changelogFragment.Page = changelogPage;

                // вставка удостоверяющего листа
                if (reportParameters.PageUD)
                {
                    var certifyingPage = new Page(doc);
                    certifyingPage.Rectangle = new Rectangle(0, 0, 210, 297);
                    // временная страница для удаления фрагментов
                    var tempPage2 = new Page(doc);
                    foreach (Fragment fragment in certifyingPage.GetFragments())
                    {
                        fragment.Page = tempPage2;
                    }
                    doc.DeletePage(tempPage2, new DeleteOptions(true) { DeletePageObjects = true });

                    var fileReferenceCertifying = new FileReference();
                    var fileObjectCertifying = fileReferenceCertifying.FindByRelativePath(
                        @"Служебные файлы\Шаблоны отчётов\Таблица проверки монтажа\Удостоверяющий лист.grb");
                    fileObjectCertifying.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
                    var сertifyingFragment = new Fragment(new FileLink(doc, fileObjectCertifying.LocalPath, true));
                    сertifyingFragment.Page = certifyingPage;
                   
                }

                var pages = new List<Page>();
                foreach (Page page in doc.GetPages())
                {
                    pages.Add(page);
                }

                for (int i = 0; i < pages.Count; i++)
                {
                    Page page = pages[i];

                    foreach (Fragment fragment in page.GetFragments())
                    {
                        foreach (FragmentVariableValue variable in fragment.GetVariables())
                        {
                            // Удаление связи с переменной документа
                            variable.AttachedVariable = null;
                            switch (variable.Name)
                            {
                                case "$denotation":
                                        variable.TextValue = reportParameters.DenotationDevice;
                                    break;
                                case "$oboznach":
                                    variable.TextValue = reportParameters.DenotationDevice;
                                    break;
                                case "$signHeader1":
                                        variable.TextValue = reportParameters.SignHeader1;
                                    break;
                                case "$signName1":
                                    variable.TextValue = reportParameters.SignName1;
                                    break;
                                case "$signHeader2":
                                    variable.TextValue = reportParameters.SignHeader2;
                                    break;
                                case "$signName2":
                                    variable.TextValue = reportParameters.SignName2;
                                    break;
                                case "$author":
                                    variable.TextValue = reportParameters.AuthorBy;
                                    break;
                                case "$musteredBy":
                                    variable.TextValue = reportParameters.MusteredBy;
                                    break;
                                case "$nControlBy":
                                    variable.TextValue = reportParameters.NControlBy;
                                    break;
                                case "$approvedBy":
                                    variable.TextValue = reportParameters.ApprovedBy;
                                    break;
                                case "$authorDate":
                                    if(reportParameters.AuthorCheck)
                                    variable.TextValue = reportParameters.AuthorByDate.ToShortDateString();
                                    break;
                                case "$musteredByDate":
                                    if (reportParameters.MusteredBCheck)
                                        variable.TextValue = reportParameters.MusteredByDate.ToShortDateString();
                                    break;
                                case "$nControlByDate":
                                    if (reportParameters.NControlByCheck)
                                        variable.TextValue = reportParameters.NControlByDate.ToShortDateString();
                                    break;
                                case "$approvedByDate":
                                    if (reportParameters.ApprovedByCheck)
                                        variable.TextValue = reportParameters.ApprovedByDate.ToShortDateString();
                                    break;
                                case "$numberList":
                                        variable.TextValue = (i+1).ToString();
                                    break;
                                case "$list":
                                    variable.TextValue = (i + 1).ToString();
                                    break;
                                case "$listov":
                                    variable.TextValue = pages.Count.ToString();
                                    break;
                                case "$firstUse":
                                    variable.TextValue = reportParameters.FirstUse;
                                    break;
                                case "$litera1":
                                    variable.TextValue = reportParameters.Litera1.ToString();
                                    break;
                                case "$litera2":
                                    variable.TextValue = reportParameters.Litera2.ToString();
                                    break;
                                case "$litera3":
                                    variable.TextValue = reportParameters.Litera3.ToString();
                                    break;
                                case "$numberSol":
                                    variable.TextValue = reportParameters.NumberSol;
                                    break;
                                case "$code":
                                    variable.TextValue = reportParameters.Code;
                                    break;
                                case "$dateSign1":
                                    if (reportParameters.DateStamp1)
                                        variable.TextValue = reportParameters.SignData1.ToShortDateString();
                                    break;
                                case "$dateSign2":
                                    if (reportParameters.DateStamp2)
                                        variable.TextValue = reportParameters.SignData2.ToShortDateString();
                                    break;
                            }
                        }
                    }
                }

                doc.EndChanges();        
                doc.Save();
                doc.Close();
            }

            private void FillSection(CadTable table, string section, List<TableInstallation> items)
            {
                var row = table.CreateRow();

                table.CreateRow();



                foreach (var item in items)
                {

                    foreach (var chain in item.ChainData)
                    {
                        row.AddText(chain.PosDenotation);
                        row.AddText(chain.StepX.ToString());
                        row.AddText(chain.StepY.ToString());
                        table.CreateRow();
                    }
                   
                }

            }
        }
    }
}

