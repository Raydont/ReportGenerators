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
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;

namespace CertificateSheetReport
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
                CertificateReport report = new CertificateReport();
                report.Make(context);

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

        public class CertificateReport : IDisposable
        {
            public IndicationForm m_form = new IndicationForm();

            public void Dispose()
            {
            }

            public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
            {
                
            }

            private IReportGenerationContext _context;

            public void Make(IReportGenerationContext context)
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
                    using (CertificateReport report = new CertificateReport())
                    {
                        m_form.setStage(IndicationForm.Stages.DataAcquisition);
                        List<TableInstallation> tableInstallations = new List<TableInstallation>();
                      
                        m_form.setStage(IndicationForm.Stages.ReportGenerating);
                        m_writeToLog("=== Формирование отчета ===");

                        var nomObject = GetNomenclatureObject();

                        //Формирование отчета
                        MakeSelectionByTypeReport(nomObject, m_writeToLog);

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

            public NomenclatureReferenceObject GetNomenclatureObject()
            {
                var referenceInfo = ReferenceCatalog.FindReference(_context.ObjectsInfo[0].ReferenceId);
                var reference = referenceInfo.CreateReference();

                var filter = new Filter(referenceInfo);
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, _context.ObjectsInfo[0].ObjectId);

                var selectedObj = reference.Find(filter).FirstOrDefault();
                return (NomenclatureReferenceObject)selectedObj;
            }

            public void MakeSelectionByTypeReport(NomenclatureReferenceObject nomObject, LogDelegate logDelegate)
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

                foreach (Page page in doc.Pages)
                {
                    page.Rectangle = new Rectangle(0, 0, 210, 297);
                    var fileReferenceCertifying = new FileReference();
                    var fileObjectCertifying = fileReferenceCertifying.FindByRelativePath(
                         @"Служебные файлы\Шаблоны отчётов\Удостоверяющий лист ДЭ\UDEKD.grb");
                    fileObjectCertifying.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
                    var сertifyingFragment = new Fragment(new FileLink(doc, fileObjectCertifying.LocalPath, true));
                    сertifyingFragment.Page = page;
                    break;
                }

                var pages = new List<Page>();
                foreach (Page page in doc.GetPages())
                {
                    pages.Add(page);
                }
                //Разраб.
                var authorSignature = nomObject.Signatures.Where(t => t.SignatureObjectType.ToString() == "Разраб.").FirstOrDefault();
                //Пров.
                var approvedSignature = nomObject.Signatures.Where(t => t.SignatureObjectType.ToString() == "Пров.").FirstOrDefault();
                //Утв.
                var musteredSignature = nomObject.Signatures.Where(t => t.SignatureObjectType.ToString() == "Утв.").FirstOrDefault();
                //Подпись ЗГКЗ           
                var zgkzSignature = nomObject.Signatures.Where(t => t.SignatureObjectType.ToString() == "Подпись ЗГКЗ").FirstOrDefault();
                //Пре.заказ.
                var vpmpSignature = nomObject.Signatures.Where(t => t.SignatureObjectType.ToString() == "Пре.заказ.").FirstOrDefault();
                //Н. контр.
                var controlSignature = nomObject.Signatures.Where(t => t.SignatureObjectType.ToString() == "Н. контр.").FirstOrDefault();
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
                                        variable.TextValue = nomObject[Guids.NomenclatureParameters.Denotation].Value.ToString();
                                    break;
                                case "$name":
                                    variable.TextValue = nomObject[Guids.NomenclatureParameters.Name].Value.ToString();
                                    break;
                                case "$nameDoc":
                                    //   variable.TextValue = nomObject.GetObjects(Guids.DocToFileLink).OrderBy(t=>t.SystemFields.CreationDate).First().ToString();
                                    break;
                                case "$author":
                                    if (authorSignature != null)
                                    {
                                        variable.TextValue = " " + authorSignature.UserObject[Guids.UserShortName].Value.ToString();
                                    }
                                    break;
                                case "$authorDate":
                                    if (authorSignature != null)
                                        variable.TextValue = ((DateTime)authorSignature.SignatureDate).ToShortDateString();
                                    break;
                                case "$musteredBy":
                                    if (musteredSignature != null)
                                        variable.TextValue = " " + musteredSignature.UserObject[Guids.UserShortName].Value.ToString();
                                    break;
                                case "$musteredByDate":
                                    if (musteredSignature != null)
                                        variable.TextValue = ((DateTime)musteredSignature.SignatureDate).ToShortDateString();
                                    break;
                                case "$nControlBy":
                                    if (controlSignature != null)
                                        variable.TextValue = " " + controlSignature.UserObject[Guids.UserShortName].Value.ToString();
                                    break;
                                case "$nControlByDate":
                                    if (controlSignature != null)
                                        variable.TextValue = ((DateTime)musteredSignature.SignatureDate).ToShortDateString();
                                    break;
                                case "$approvedBy":
                                    if (approvedSignature != null)
                                        variable.TextValue = " " + approvedSignature.UserObject[Guids.UserShortName].Value.ToString();
                                    break;
                                case "$approvedByDate":
                                    if (approvedSignature != null)
                                       variable.TextValue = ((DateTime)approvedSignature.SignatureDate).ToShortDateString();
                                    break;                          
                                case "$list":
                                    variable.TextValue = (i + 1).ToString();
                                    break;
                                case "$listov":
                                    variable.TextValue = pages.Count.ToString();
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

