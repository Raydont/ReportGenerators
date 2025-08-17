using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;

using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.References;
using ReportHelpers;

using Microsoft.Office.Interop.Excel;
using InteropExcel = Microsoft.Office.Interop.Excel;

namespace DocumentSendingSummaryReport
{
    public class DocumentSendingSummaryReport : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xls";
            }
        }

        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        public void Generate(IReportGenerationContext context)
        {
            if (File.Exists(context.ReportFilePath))
            {
                try
                {
                    File.Delete(context.ReportFilePath);
                }
                catch
                {
                    MessageBox.Show("Закройте файл " + context.ReportFilePath + " перед формированием отчета!");
                    return;
                }
            }

            context.CopyTemplateFile();    // Создаем копию шаблона            

            // Получение ID объекта, на который получаем отчет
            int baseDocumentID = Initialize(context);
            if (baseDocumentID == -1) return;

            // Справочник РКК
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(TFDClass.regCardsGuid);
            var regCardsReference = referenceInfo.CreateReference();

            var document = regCardsReference.Find(baseDocumentID);
            if ((document.Class.Guid != TFDClass.inputDocGuid) && (document.Class.Guid != TFDClass.orderDocGuid) && (document.Class.Guid != TFDClass.raspDocGuid))
            {
                MessageBox.Show("Документ данного типа в рассылке не задействован", "ВНИМАНИЕ!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                // Формирование отчета
                Xls xls = new Xls();
                bool needMake = false;

                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(context.ReportFilePath);
                    needMake = MakeReport(document, baseDocumentID, xls);
                }
                finally
                {
                    xls.Quit(true);
                    if (needMake)
                        Process.Start(context.ReportFilePath);
                }
            }
        }

        public int CreateTableHeader(string header, int row, Xls xls)
        {
            xls[2, row].SetValue(header);
            xls[2, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[2, row].Font.Bold = true;
            xls[2, row].Font.Underline = true;
            xls[1, row, 4, 1].Cells.MergeCells = true;
            row++;
            return row;
        }

        public int CreateHeader(int row, Xls xls)
        {
            xls[1, row].SetValue("Отправитель");
            xls[1, row].Font.Bold = true;
            xls[1, row].Interior.Color = Color.SteelBlue;
            xls[1, row].Font.Color = Color.White;

            xls[2, row].SetValue("Получатель");
            xls[2, row].Font.Bold = true;
            xls[2, row].Interior.Color = Color.SteelBlue;
            xls[2, row].Font.Color = Color.White;

            xls[3, row].SetValue("Задание");
            xls[3, row].Font.Bold = true;
            xls[3, row].Interior.Color = Color.SteelBlue;
            xls[3, row].Font.Color = Color.White;

            xls[4, row].SetValue("Срок");
            xls[4, row].Font.Bold = true;
            xls[4, row].Interior.Color = Color.SteelBlue;
            xls[4, row].Font.Color = Color.White;

            xls[1, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            row++;
            return row;
        }

        public void CreateBorderTable(int row, int rowNach, Xls xls)
        {
            xls[1, rowNach, 4, row - rowNach].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, rowNach, 4, row - rowNach].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            xls[1, rowNach, 4, 1].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
        }

        public int FillDataSet(ReferenceObject distrib, int row, Xls xls)
        {
            ReferenceObjectCollection distribs = distrib.Links.ToMany[TFDClass.distributionObjectsListGuid].Objects;

            foreach (var distr in distribs)
            {
                ReferenceObjectCollection receivers = distr.Links.ToMany[TFDClass.distributionObjectReceiversLinkGuid].Objects;
                foreach (var receiver in receivers)
                {
                    xls[1, row].SetValue(distrib[TFDClass.senderDocumentGuid].Value.ToString());
                    xls[1, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[1, row].Interior.Color = Color.AntiqueWhite;
                    xls[2, row].SetValue(receiver.ToString());
                    xls[2, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[2, row].Interior.Color = Color.LightCyan;
                    xls[3, row].SetValue(distr[TFDClass.distributionObjectCommentGuid].Value.ToString());
                    xls[3, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[3, row].Interior.Color = Color.AntiqueWhite;
                    xls[4, row].SetValue(((DateTime)distr[TFDClass.distributionObjectDeadlineGuid].Value).ToShortDateString());
                    xls[4, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[4, row].Interior.Color = Color.LightCyan;
                    row++;
                }
            }
            return row;
        }

        public int CreateParameterTable(ReferenceObject document, Xls xls)
        {
            // Создание таблицы параметров документа
            xls[2, 1].Font.Size = 12;
            xls[2, 1].Font.Bold = true;
            xls[2, 1].SetValue("Рассылка документа " + document[TFDClass.regNumGuid].Value + " от " + ((DateTime)document[TFDClass.regDateGuid].Value).ToShortDateString());
            xls[2, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1, 4, 1].Cells.MergeCells = true;

            xls[1, 3].SetValue("Параметр");
            xls[1, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 3].Font.Bold = true;
            xls[1, 3].Interior.Color = Color.SteelBlue;
            xls[1, 3].Font.Color = Color.White;
            xls[2, 3].SetValue("Значение параметра");
            xls[2, 3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[2, 3].Font.Bold = true;
            xls[2, 3].Interior.Color = Color.SteelBlue;
            xls[2, 3].Font.Color = Color.White;
            xls[2, 3, 3, 1].Cells.MergeCells = true;

            // Ширина столбцов
            xls[1, 3].ColumnWidth = 40;
            xls[2, 3].ColumnWidth = 40;
            xls[3, 3].ColumnWidth = 50;
            xls[4, 3].ColumnWidth = 10;


            int row;

            // Если это Входящий документ, то создаем таблицу его параметров
            if (document.Class.Guid == TFDClass.inputDocGuid)
            {
                xls[1, 4].SetValue("Исходящий номер");
                xls[1, 5].SetValue("Дата исходящего");
                xls[1, 6].SetValue("Отправитель");
                xls[1, 7].SetValue("Адрес отправителя");
                xls[1, 8].SetValue("Краткое содержание");

                xls[2, 4].SetValue(document[TFDClass.incomingIndexGuid].Value.ToString());
                xls[2, 5].SetValue(((DateTime)document[TFDClass.incomingDateGuid].Value).ToShortDateString());
                xls[2, 6].SetValue(document[TFDClass.organisationGuid].Value.ToString());
                xls[2, 7].SetValue(document[TFDClass.cityGuid].Value.ToString());

                row = 8;

                foreach (string comment in document[TFDClass.summaryIndex].Value.ToString().Split('\n'))
                {
                    xls[2, row].SetValue(comment);
                    row++;
                }
            }
            else // Для Приказов по предприятию и Распоряжений по предприятию
            {
                xls[1, 4].SetValue("Краткое содержание");
                
                row = 4;
                
                foreach (string comment in document[TFDClass.contentGuid].Value.ToString().Split('\n'))
                {
                    xls[2, row].SetValue(comment);
                    row++;
                }

                xls[1, row].SetValue("Документ подписал");

                foreach (var user in document.Links.ToMany[TFDClass.RCUsersLinkGuid].Objects)
                {
                    xls[2, row].SetValue(user.ToString());
                    row++;
                }
            }

            for (int i = 4; i <= row; i++)
            {
                xls[2, i, 3, 1].Cells.MergeCells = true;
                xls[1, i - 1, 4, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            }

            xls[1, 4, 1, row - 4].Interior.Color = Color.AntiqueWhite;
            xls[2, 4, 3, row - 4].Interior.Color = Color.LightCyan;
            xls[1, 3, 4, 1].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            xls[1, 4, 4, row - 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

            row++;
            return row;
        }


        public bool MakeReport(ReferenceObject document, int baseDocumentID, Xls xls)
        {
            bool needMake;
            // Справочник "Группы и пользователи"
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(TFDClass.userReferenceGuid);
            var userReference = referenceInfo.CreateReference();
            // Поиск пользователей, входящих в группу "Канцелярия"
            ReferenceObject chancellory = userReference.Find(TFDClass.chancellory920Guid);
            ReferenceObjectCollection chancelloryUsers = chancellory.Children;

            // Справочник "Списки рассылки"
            referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(TFDClass.distributionListGuid);
            var distributionListReference = referenceInfo.CreateReference();

            // Поиск элементов справочника "Списки рассылки", к которым привязана выбранная РКК, а отправитель - сотрудник(и) канцелярии
            var filter = new TFlex.DOCs.Model.Search.Filter(referenceInfo);
            filter.Terms.AddTerm("[" + TFDClass.connectedRKKGuid + "]->ID", ComparisonOperator.Equal, baseDocumentID);
            filter.Terms.AddTerm("Автор", ComparisonOperator.IsOneOf, chancelloryUsers);
            List<ReferenceObject> distributionList = distributionListReference.Find(filter);

            if (distributionList.Count > 0)
            {
                needMake = true;
                // Таблица параметров документа
                int row = CreateParameterTable(document, xls);
                int rowNach;

                foreach (var distrib in distributionList)
                {
                    row = CreateTableHeader("Рассылка из канцелярии", row, xls);
                    row++;

                    // Заголовок
                    rowNach = row;
                    row = CreateHeader(row, xls);
                    row = FillDataSet(distrib, row, xls);
                    CreateBorderTable(row, rowNach, xls);
                    row++;
                }

                ReferenceObjectCollection distributions;

                while (distributionList.Count > 0)
                {
                    // Загружаем рассылку, одновременно удаляем из списка рассылок
                    ReferenceObject distribution = distributionList[0];
                    distributionList.Remove(distribution);
                    // Пользователи или подразделения, которым направлен документ
                    distributions = distribution.Links.ToMany[TFDClass.distributionObjectsListGuid].Objects.FirstOrDefault().Links.ToMany[TFDClass.distributionObjectReceiversLinkGuid].Objects;                    

                    foreach (var distributor in distributions)
                    {
                        List<ReferenceObject> findResults = new List<ReferenceObject>();

                        // Если distributor - не пользователь, а подразделение, то ищем его начальника(ов)
                        if (distributor.Class.IsInherit(TFDClass.departmentClassGuid))
                        {
                            var bossGroup = distributor.Children.Where(t => t[TFDClass.usersNameGuid].GetString() == "Начальник").FirstOrDefault();

                            if (bossGroup != null)
                            {
                                ReferenceObjectCollection bosses = bossGroup.Children;

                                foreach (var boss in bosses)
                                {
                                    // Если руководитель отдела-получателя не является отправителем, то обрабатываем, иначе пропускаем (во избежание зацикливания)
                                    if (boss.ToString() != distribution[TFDClass.senderDocumentGuid].Value.ToString())
                                    {
                                        // Поиск элементов справочника "Списки рассылки", к которым привязана выбранная РКК, а отправитель - начальник(и) указанного подразделения
                                        filter.Terms.Clear();
                                        filter.Terms.AddTerm("[" + TFDClass.connectedRKKGuid + "]->ID", ComparisonOperator.Equal, baseDocumentID);
                                        filter.Terms.AddTerm("Автор", ComparisonOperator.Equal, boss);

                                        findResults = distributionListReference.Find(filter);

                                        if (findResults.Count > 0)
                                        {
                                            // Добавление элементов в список рассылок
                                            distributionList.AddRange(findResults);

                                            row = CreateTableHeader("Рассылка подразделения " + distributor.ToString(), row, xls);
                                            row++;

                                            // Заголовок
                                            rowNach = row;
                                            row = CreateHeader(row, xls);

                                            foreach (var distr in findResults)
                                            {
                                                row = FillDataSet(distr, row, xls);
                                            }
                                            CreateBorderTable(row, rowNach, xls);
                                            row++;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Если адресат рассылки не равен автору рассылки, то работаем, иначе пропускаем (во избежание зацикливания)
                            if (distributor.ToString() != distribution[TFDClass.senderDocumentGuid].Value.ToString())
                            {
                                // Поиск элементов справочника "Списки рассылки", к которым привязана выбранная РКК, а отправитель - указанный пользователь
                                filter.Terms.Clear();
                                filter.Terms.AddTerm("[" + TFDClass.connectedRKKGuid + "]->ID", ComparisonOperator.Equal, baseDocumentID);
                                filter.Terms.AddTerm("Автор", ComparisonOperator.Equal, distributor);

                                findResults = distributionListReference.Find(filter);

                                if (findResults.Count > 0)
                                {
                                    // Добавление элементов в список рассылок
                                    distributionList.AddRange(findResults);

                                    row = CreateTableHeader("Рассылка пользователя " + distributor.ToString(), row, xls);
                                    row++;

                                    // Заголовок
                                    rowNach = row;
                                    row = CreateHeader(row, xls);

                                    foreach (var distr in findResults)
                                    {
                                        row = FillDataSet(distr, row, xls);
                                    }
                                    CreateBorderTable(row, rowNach, xls);
                                    row++;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Список рассылки документа не заполнен", "ВНИМАНИЕ!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                needMake = false;
            }

            return needMake;
        }

        static int Initialize(IReportGenerationContext context)
        {
            // --------------------------------------------------------------------------
            // Инициализация
            // --------------------------------------------------------------------------

            int documentID;

            // Получаем ID выделенного в интерфейсе T-FLEX DOCs объекта
            if (context.ObjectsInfo.Count == 1) documentID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }

            return documentID;
        }
    }
}
