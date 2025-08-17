using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex;
using System.Drawing;

namespace AccountingCardReport
{
    public class ReportGenerator : IReportGenerator
    {
        /// <summary>
        /// Расширение файла отчета
        /// </summary>
        public string DefaultReportFileExtension
        {
            get
            {
                return ".grb";
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
            try
            {
                var refID = context.ReferenceId;
                var currentObjectID = context.ObjectsInfo[0].ObjectId;
                ReferenceObject currentObject;
                DataReader reader;
                DocumentParameters documentParameters;
                // Класс, форматирующий текст
                TextFormatter tForm = new TextFormatter(new System.Drawing.Font("T-FLEX Type T", 2.0f, GraphicsUnit.Millimeter));

                switch (refID)
                {
                    case Settings.nomID:
                    {
                        var nomReference = new NomenclatureReference(ServerGateway.Connection);
                        // Текущий объект номенклатуры
                        currentObject = nomReference.Find(currentObjectID);
                        // Чтение данных
                        reader = new DataReader(nomReference, tForm);
                        documentParameters = reader.GetDocumentData(currentObject);
                        // Инициализация CAD APi
                        var cadDocument = GetCadDocument(context);
                        if (cadDocument == null) MessageBox.Show("Нет документа");
                        var maker = new ReportMaker(cadDocument);
                        maker.FillDocument(documentParameters);
                        cadDocument.Save();
                        cadDocument.Close();
                        System.Diagnostics.Process.Start(context.ReportFilePath);
                        break;
                    }
                    case Settings.stoID:
                    {
                        // Справочник СТО
                        var stoRefInfo = ServerGateway.Connection.ReferenceCatalog.Find(Settings.stoID);
                        var stoReference = stoRefInfo.CreateReference();
                        // Текущий объект СТО
                        currentObject = stoReference.Find(currentObjectID);
                        // Чтение данных
                        reader = new DataReader(tForm);
                        documentParameters = reader.GetSTODocumentData(currentObject);
                        // Инициализация CAD APi
                        var cadDocument = GetCadDocument(context);
                        if (cadDocument == null) MessageBox.Show("Нет документа");
                        var maker = new ReportMaker(cadDocument);
                        maker.FillDocument(documentParameters);
                        cadDocument.Save();
                        cadDocument.Close();
                        System.Diagnostics.Process.Start(context.ReportFilePath);
                        break;
                    }
                }
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.ToString(), "Ошибка!"); 
            }
        }

        /// <summary>
        /// Инициализация CAD APi
        /// </summary>
        public TFlex.Model.Document GetCadDocument(IReportGenerationContext context)
        {
            var setup = new ApplicationSessionSetup
            {
                Enable3D = true,
                ReadOnly = false,
                PromptToSaveModifiedDocuments = false,
                EnableMacros = true,
                EnableDOCs = true,
                DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version13,
                ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
            };
            bool result = TFlex.Application.InitSession(setup);

            if (!result)
                throw new InvalidOperationException("Ошибка подключения к Cad Api");

            TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.DoNotRefresh;     //Инициализация API CAD
            context.CopyTemplateFile();
            return TFlex.Application.OpenDocument(context.ReportFilePath);
        }
    }
  
    public class DocumentParameters
    {
        public string name;                               // Наименование
        public string denotation;                         // Обозначение
        public string doctype;                            // Вид документа (бумага, электрография)
        public string department;                         // Подразделение
        public string invnum;                             // Инвентарный номер (текущий)
        public string manufactoryCode;                    // Код предприятия
        public string original;                           // Подлинник на предприятии
        public string sheetsCount;                        // Количество листов
        public string incomingDate;                       // Дата поступления
        public string format;                             // Формат документа
        // Набор строк таблицы Формы 2
        public Dictionary<int, List<string>> tableForm2 = new Dictionary<int, List<string>>();
        // Набор строк таблицы Формы 2а
        public Dictionary<int, List<string>> tableForm2a = new Dictionary<int, List<string>>();
    }

    // Класс "Использование документа"
    public class DocumentUseString
    {
        public int makeNumber;                                // Номер исполнения
        public DateTime date;                                 // Дата включения в состав
        public int useLevel;                                  // Уровень применяемости: 0 - полная первичная, 1 - прямая первичная, 2 - без первичной
        public string directUse;                              // Прямая входимость (обозначение)
        public string complexUse;                             // Входимость в комплекс
        public List<string> useList = new List<string>();     // Прямая входимость + входимость в комплекс (после переноса)
    }
    // Класс "Изменение документа"
    public class DocumentMod
    {
        public int modNumber;                                 // Номер изменения
        public string denotation;                             // Обозначение документа-извещения
        public DateTime date;                                 // Дата извещения
        public string content;                                // Содержание изменения (изменяемые листы и т.д.)
        public List<string> contentList = new List<string>(); // Содержание изменения (после переноса)
    }
    // Класс "Абонент" - общий для внешнего и внутреннего
    public class Customer
    {
        public string makeNumber;                             // Номер исполнения
        public string customer;                               // Абонент
        public DateTime date;                                 // Дата выдачи
        public int count;                                     // Количество копий
        public string based;                                  // Документ-основание
        public bool cancelled;                                // Аннулировано
    }
}
