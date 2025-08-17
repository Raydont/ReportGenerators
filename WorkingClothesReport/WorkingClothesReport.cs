using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.Search;
using ReportHelpers;
using Microsoft.Office.Interop.Excel;

using Filter = TFlex.DOCs.Model.Search.Filter;

namespace WorkingClothesReport
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
            // Проверка, существует ли файл по указанному пути (для всех типов отчета)
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

            try
            {
                // Справочник "Пользователи"
                var UserReferenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(TFDClass.usersGuid);
                var UserReference = UserReferenceInfo.CreateReference();
                // Получение всех подразделений АО "РКБ Глобус"
                // Поиск ОАО "РКБ Глобус" по ID из справочника "Группы и пользователи"
                var globus = UserReference.Find(87);
                // Все входящие в АО "РКБ Глобус", где тип объектов наследуется от "Производственное подразделение", и не являющиеся руководством предприятия
                List<ReferenceObject> departments = globus.Children.Where(t => t.Class.IsInherit(TFDClass.manufactoryDepartment) && t.SystemFields.Id != 687).ToList();

                ReportForm reportForm = new ReportForm();
                reportForm.departments = departments.OrderBy(t => t.ToString()).ToList();
                reportForm.MakeReport(context);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
    
    public class WorkingClothesReport : IDisposable
    {
        public void Dispose()
        {
        }

        public bool Make(bool isOverall, string depart, string halfYear, string year, Xls xls)
        {
            bool needMake = true;

            // Справочник "Закупки"
            var GPurchasesReferenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(TFDClass.globPurchasesGuid);
            var GPurchaseReference = GPurchasesReferenceInfo.CreateReference();

            // Выбираем элемент, где тип - "Закупка", а название состоит из года и полугодия, выбранных в диалоге
            Filter filter = new Filter(GPurchasesReferenceInfo);
            filter.Terms.AddTerm("Тип", ComparisonOperator.Equal, GPurchaseReference.Classes.Find(TFDClass.purchaseTypeGuid));
            filter.Terms.AddTerm("[Наименование]", ComparisonOperator.Equal, year + "/" + halfYear);

            ReferenceObject purchase = GPurchaseReference.Find(filter).FirstOrDefault();
            // Если закупка не найдена, то отчет не формируется
            if (purchase == null)
            {
                MessageBox.Show("Закупок за данный период не зарегистрировано", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            
            // Поиск заявок, входящих в закупку по соответствующей связи
            List<ReferenceObject> clothesClaims = new List<ReferenceObject>();
            if (isOverall)
                // Для сводного отчета забираем все входящие в закупку заявки
                clothesClaims = purchase.GetObjects(TFDClass.claimsLinkGuid);
            else
            {
                // Справочник "Заявки на спецодежду"
                var GClaimsReferenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(TFDClass.globClaimsGuid);
                var GClaimsReference = GClaimsReferenceInfo.CreateReference();
                // Очистка фильтра от предыдущих условий
                filter = new Filter(GClaimsReferenceInfo);
                // Новые условия в справочнике "Заявки на спецодежду"
                filter.Terms.AddTerm("[Закупка]->[Наименование]", ComparisonOperator.Equal, year + "/" + halfYear);
                filter.Terms.AddTerm("[Получатель спецодежды]->[Подразделение]->[Наименование]", ComparisonOperator.Equal, depart);
                // Список заявок выбранного подразделения
                clothesClaims = GClaimsReference.Find(filter).ToList();
            }
            // Если заявок нет, то отчет не формируется
            if (clothesClaims.Count == 0)
            {
                MessageBox.Show("Ни одной заявки на данный период не зарегистрировано", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Список объектов класса "Заявка на спецодежду"
            var claimsList = ReadData(clothesClaims);

            // Настройки страницы
            PageSetup(isOverall, xls);
            // Шапка с подписями
            CreateTopSignBlock(isOverall, xls);
            // Название отчета
            CreateReportName(isOverall, depart, halfYear, year, xls);
            // Заголовок
            MakeHeader(isOverall, xls);
            // Таблица отчета
            int row = MakeReportBody(isOverall, claimsList, depart, xls);
            // Если row = 0, что означает, что заявок от подразделения нет, то отчет не формируется
            if (row == 0)
            {
                MessageBox.Show("Ни одной заявки от данного подразделения не зарегистрировано", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Нижний блок подписей
            CreateBottomSignBlock(row, isOverall, xls);

            return needMake;
        }

        // Чтение данных из списка заявок на закупку
        public List<ClothesClaim> ReadData(List<ReferenceObject> claims)
        {
            ProgressForm p_form = new ProgressForm();
            p_form.Text = "Идет сбор данных...";
            p_form.Visible = true;

            p_form.max = claims.Count;
            p_form.setProgressBarMax();
            int step = 1; 

            List<ClothesClaim> result = new List<ClothesClaim>();

            // Справочник "Заявки на спецодежду"
            var GClaimsReferenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(TFDClass.globClaimsGuid);
            var GClaimsReference = GClaimsReferenceInfo.CreateReference();
            {
                GClaimsReference.LoadSettings.Clear();
                GClaimsReference.LoadSettings.AddParameters(TFDClass.clothesSizeGuid, TFDClass.clothesHeightGuid, TFDClass.clothesMaterialGuid, TFDClass.clothesColorGuid, TFDClass.clothesTechParamsGuid,
                    TFDClass.clothesCountGuid, TFDClass.clothesDateGuid, TFDClass.commentGuid);
                var reqUserRelation = GClaimsReference.LoadSettings.AddRelation(TFDClass.reqUserLinkGuid);
                reqUserRelation.AddParameters(TFDClass.reqUserFIOGuid, TFDClass.reqUserFIOGuid, TFDClass.reqUserProfessionGuid, TFDClass.reqUserGenderGuid);
                var depRelation = reqUserRelation.AddRelation(TFDClass.departmentLinkGuid);
                var normRelation = GClaimsReference.LoadSettings.AddRelation(TFDClass.clothesTypeLinkGuid);
                normRelation.AddParameters(TFDClass.measureUnitGuid, TFDClass.clothesNameGuid);
            }
            GClaimsReference.LoadLinks(claims, GClaimsReference.LoadSettings);

            foreach (var claim in claims)
            {
                var element = new ClothesClaim();

                element.ClothesSize = claim[TFDClass.clothesSizeGuid].GetString();
                element.ClothesHeight = claim[TFDClass.clothesHeightGuid].GetString();
                element.ClothesMaterial = claim[TFDClass.clothesMaterialGuid].GetString();
                element.ClothesColor = claim[TFDClass.clothesColorGuid].GetString();
                element.ClothesTechParams = claim[TFDClass.clothesTechParamsGuid].GetString();
                element.ClothesCount = claim[TFDClass.clothesCountGuid].GetInt32();
                element.ClothesDate = claim[TFDClass.clothesDateGuid].GetDateTime();
                element.Comment = claim[TFDClass.commentGuid].GetString();

                // Связь "Получатель"
                element.ReqUserFIO = claim.GetObjectValue("[Получатель спецодежды]->[ФИО]").ToString();
                element.ReqUserProfession = claim.GetObjectValue("[Получатель спецодежды]->[Профессия]").ToString();
                element.ReqUserGender = claim.GetObjectValue("[Получатель спецодежды]->[Пол]").ToString();
                // Связь "Подразделение"
                element.Department = claim.GetObjectValue("[Получатель спецодежды]->[Подразделение]->[Наименование]").ToString();
                // Связь "Норматив из справочника"
                element.MeasureUnit = claim.GetObjectValue("[Норматив из справочника]->[Единица измерения]").ToString();
                element.ClothesName = claim.GetObjectValue("[Норматив из справочника]->[Наименование СИЗ]").ToString(); ;

                // Параметр для группировки
                element.UnionParameter = (element.ClothesName + "#" + element.ClothesSize + "#" + element.ClothesHeight + "#" 
                                        + element.ClothesMaterial + "#" + element.ClothesColor + "#" + element.ReqUserGender).Trim();

                result.Add(element);
                p_form.progressBarParam(step);
                System.Windows.Forms.Application.DoEvents();
            }

            p_form.Close();

            return result;
        }

        // Заполнение таблицы отчета
        private int MakeReportBody(bool isOverall, List<ClothesClaim> claims, string department, Xls xls)
        {
            int row;
            int cnt = 1;

            if (isOverall)
            {
                row = 9;
                // Формирование сводной ведомости заказа спецодежды

                // продистинктенный (по названию, размеру, росту, материалу, цвету и полу) список спецодежды
                IEqualityComparer<ClothesClaim> comparer = new UnionParameterComparer();
                var distinct = claims.Distinct(comparer).OrderBy(t => t.ClothesName).ThenBy(t => t.ReqUserGender).ThenBy(t => t.ClothesSize).
                                                                                              ThenBy(t=> t.ClothesHeight).ThenBy(t => t.ClothesMaterial).ThenBy(t => t.ClothesColor).ToList();

                ProgressForm p_form = new ProgressForm();
                p_form.Text = "Идет формирование отчета...";
                p_form.Visible = true;

                p_form.max = distinct.Count;
                p_form.setProgressBarMax();
                int step = 1; 

                foreach (var item in distinct)
                {
                    // Список всех элементов входного массива, у которого Параметр для группировки тот же, что и в item
                    var itemMass = claims.Where(t => t.UnionParameter == item.UnionParameter).ToList();

                    xls[1, row].Value = cnt;
                    xls[2, row].Value = item.ClothesName;
                    xls[3, row].Value = itemMass.Sum(t => t.ClothesCount);          // Общее количество
                    xls[4, row].Value = item.ClothesSize;
                    xls[5, row].NumberFormat = "@";
                    xls[5, row].Value = item.ClothesHeight;
                    xls[6, row].Value = item.MeasureUnit;
                    xls[7, row].Value = item.ClothesMaterial;
                    xls[8, row].Value = item.ClothesColor;
                    xls[9, row].Value = (item.ReqUserGender == "0") ? "Мужской" : "Женский";

                    // Список подразделений, заказавших данный лот
                    var departs = itemMass.Select(t => t.Department).Distinct().ToList();

                    StringBuilder sb = new StringBuilder();
                    foreach (var dep in departs)
                    {
                        // Суммирование по каждому подразделению
                        sb.AppendLine(dep + "(" + itemMass.Where(t => t.Department == dep).Sum(t => t.ClothesCount) + ")");
                    }

                    xls[10, row].Value = sb.ToString();
                    xls[1, row, 10, 1].WrapText = true;
                    xls[1, row, 10, 1].VerticalAlignment = XlVAlign.xlVAlignTop;

                    row++;
                    cnt++;

                    p_form.progressBarParam(step);
                    System.Windows.Forms.Application.DoEvents();
                }

                xls[1, 9, 10, row - 9].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                xls[1, 9, 10, row - 9].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight);
                p_form.Close();
            }
            else
            {
                // Формирование Заявки на спецодежду подразделения
                row = 10;
                // Фильтрация по заданному подразделению
                var departmentClaims = claims.Where(t => t.Department == department).OrderBy(t => t.ReqUserFIO).ThenBy(t => t.ClothesName).ThenBy(t => t.ReqUserGender).
                                                                                     ThenBy(t => t.ClothesSize).ThenBy(t => t.ClothesHeight).ThenBy(t => t.ClothesMaterial).
                                                                                     ThenBy(t => t.ClothesColor).ToList();
                if (departmentClaims.Count == 0)
                    return 0;

                ProgressForm p_form = new ProgressForm();
                p_form.Text = "Идет формирование отчета...";
                p_form.Visible = true;

                p_form.max = departmentClaims.Count;
                p_form.setProgressBarMax();
                int step = 1; 

                foreach (var claim in departmentClaims)
                {
                    xls[1, row].Value = cnt;
                    xls[2, row].Value = claim.ClothesName;
                    xls[3, row].Value = claim.ReqUserFIO;
                    xls[4, row].Value = claim.ClothesSize;
                    xls[5, row].NumberFormat = "@";
                    xls[5, row].Value = claim.ClothesCount;
                    xls[6, row].Value = claim.ClothesColor;
                    xls[7, row].Value = claim.ClothesTechParams;
                    xls[8, row].Value = claim.Comment;
                    xls[1, row, 8, 1].WrapText = true;

                    xls[1, row, 8, 1].VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[2, row + 1].Value = claim.ClothesMaterial;
                    xls[3, row + 1].Value = claim.ReqUserProfession;
                    xls[4, row + 1].NumberFormat = "@";
                    xls[4, row + 1].Value = claim.ClothesHeight;
                    xls[5, row + 1].Value = claim.MeasureUnit;
                    xls[6, row + 1].Value = (claim.ReqUserGender == "0") ? "Мужской" : "Женский";
                    string clothesDate = claim.ClothesDate.ToString("MMMM yyyy");
                    xls[7, row + 1].Value = clothesDate == "Январь 0001" ? String.Empty : clothesDate;
                    xls[2, row + 1, 6, 1].WrapText = true;

                    xls[2, row + 1, 6, 1].VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[1, row, 1, 2].Merge();
                    xls[8, row, 1, 2].Merge();
                    xls[1, row, 8, 2].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                    xls[1, row, 8, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeTop,
                                                                                                    XlBordersIndex.xlEdgeBottom,
                                                                                                    XlBordersIndex.xlEdgeLeft,
                                                                                                    XlBordersIndex.xlEdgeRight);
                    row += 2;
                    cnt++;

                    p_form.progressBarParam(step);
                    System.Windows.Forms.Application.DoEvents();
                }

                p_form.Close();
            }

            return row;
        }

        // Создание заголовка отчета
        private void MakeHeader(bool isOverall, Xls xls)
        {
            int row = 8;

            if (isOverall)
            {
                xls[1, row].Value = "№ п/п";
                xls[2, row].Value = "Вид продукции";
                xls[3, row].Value = "Общ. кол-во";
                xls[4, row].Value = "Размер";
                xls[5, row].Value = "Рост";
                xls[6, row].Value = "Ед. изм.";
                xls[7, row].Value = "Материал";
                xls[8, row].Value = "Цвет";
                xls[9, row].Value = "Пол";
                xls[10, row].Value = "Подразделения";

                xls[1, row, 10, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[1, row, 10, 1].Font.Bold = true;
                xls[1, row, 10, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }
            else
            {
                xls[1, row].Value = "№ п/п";
                xls[2, row].Value = "Вид продукции";
                xls[3, row].Value = "Ф.И.О.";
                xls[4, row].Value = "Размер";
                xls[5, row].Value = "Кол-во";
                xls[6, row].Value = "Цвет";
                xls[7, row].Value = "Тех.х-ки";
                xls[8, row].Value = "Примечания";
                xls[2, row + 1].Value = "Материал";
                xls[3, row + 1].Value = "Профессия";
                xls[4, row + 1].Value = "Рост";
                xls[5, row + 1].Value = "Ед. изм.";
                xls[6, row + 1].Value = "Пол";
                xls[7, row + 1].Value = "Срок обесп-я";
                xls[1, row, 1, 2].Merge();
                xls[8, row, 1, 2].Merge();

                xls[1, row, 8, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[1, row, 8, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
                xls[1, row, 8, 2].Font.Bold = true;
                xls[1, row, 8, 2].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                xls[1, row, 8, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeBottom);
            }
        }

        // Шапка с подписями
        private void CreateTopSignBlock(bool isOverall, Xls xls)
        {
            xls[1, 1].Value = "  СОГЛАСОВАНО";
            xls[1, 2].Value = "Начальник бюро 924";
            xls[1, 3].Value = "__________ Коломойцева И.Н.";
            xls[1, 1, 1, 3].Font.Bold = true;

            if (isOverall)
            {
                xls[9, 1].Value = "     УТВЕРЖДАЮ";
                xls[9, 2].Value = "Зам. ген. директора";
                xls[9, 3].Value = "__________ Голыхов А.В.";
                xls[9, 1, 1, 3].Font.Bold = true;
            }
            else
            {
                xls[8, 1].Value = "     УТВЕРЖДАЮ";
                xls[8, 2].Value = "Зам. ген. директора";
                xls[8, 3].Value = "__________ Голыхов А.В.";
                xls[8, 1, 1, 3].Font.Bold = true;
            }
        }

        // Название отчета
        private void CreateReportName(bool isOverall, string depart,  string halfYear, string year, Xls xls)
        {
            if (isOverall)
            {
                xls[1, 5, 10, 1].Merge();
                xls[1, 5].Value = "Сводная ведомость заказа спецодежды";
                xls[1, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[1, 6, 10, 1].Merge();
                xls[1, 6].Value = "на " + halfYear + " полугодие " + year + " года";
                xls[1, 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }
            else
            {
                xls[1, 5, 8, 1].Merge();
                xls[1, 5].Value = "Заявка на спецодежду подразделения " + depart;
                xls[1, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[1, 6, 8, 1].Merge();
                xls[1, 6].Value = "на " + halfYear + " полугодие " + year + " года";
                xls[1, 6].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }

            xls[1, 5, 1, 2].Font.Size = 12;
            xls[1, 5, 1, 2].Font.Bold = true;
        }

        // Подписи исполнителя и его начальника
        private void CreateBottomSignBlock(int row, bool isOverall, Xls xls)
        {
            row++;
            xls[1, row, 1, 2].RowHeight = 30;
            xls[1, row].Value = "Исполнитель";
            xls[1, row + 1].Value = "Начальник подразделения";

            if (isOverall)
            {
                xls[5, row].Value = "тел.";
                xls[1, row, 5, 2].Font.Bold = true;
            }
            else
            {
                xls[4, row].Value = "тел.";
                xls[1, row, 4, 2].Font.Bold = true;
            }
        }

        // Параметры листа Excel (под формат А4)
        private void PageSetup(bool isOverall, Xls xls)
        {
            PageSetup pageSetup = xls.Worksheet.PageSetup;
            pageSetup.PaperSize = XlPaperSize.xlPaperA4;
            pageSetup.Orientation = XlPageOrientation.xlLandscape;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 500;
            pageSetup.TopMargin = 25.0;
            pageSetup.BottomMargin = 13.0;
            pageSetup.LeftMargin = 25.0;
            pageSetup.RightMargin = 13.0;
            pageSetup.Zoom = false;

            // Установка ширины строк по альбомному А4
            if (isOverall)
            {
                xls[1, 1].ColumnWidth = 6;
                xls[2, 1].ColumnWidth = 22;
                xls[3, 1].ColumnWidth = 11;
                xls[4, 1].ColumnWidth = 8;
                xls[5, 1].ColumnWidth = 8;
                xls[6, 1].ColumnWidth = 8;
                xls[7, 1].ColumnWidth = 21;
                xls[8, 1].ColumnWidth = 16;
                xls[9, 1].ColumnWidth = 9;
                xls[10, 1].ColumnWidth = 21;
            }
            else
            {
                xls[1, 1].ColumnWidth = 6;
                xls[2, 1].ColumnWidth = 24;
                xls[3, 1].ColumnWidth = 23;
                xls[4, 1].ColumnWidth = 8;
                xls[5, 1].ColumnWidth = 8;
                xls[6, 1].ColumnWidth = 18;
                xls[7, 1].ColumnWidth = 23;
                xls[8, 1].ColumnWidth = 24;
            }

            xls[1, 1, 1, 3].RowHeight = 22.5;
            xls[1, 5, 1, 2].RowHeight = 22.5;
        }
    }

    // Класс "Заявка на спецодежду"
    public class ClothesClaim
    {
        public string ClothesName;                                   // Наименование
        public string ClothesSize;                                   // Размер
        public string ClothesHeight;                                 // Рост
        public string ClothesMaterial;                               // Материал
        public string ClothesColor;                                  // Цвет
        public string ClothesTechParams;                             // Тех.х-ки
        public int ClothesCount;                                     // Количество
        public DateTime ClothesDate;                                 // Дата поставки
        public string Comment;                                       // Примечание
        public string ReqUserFIO;                                    // ФИО
        public string ReqUserProfession;                             // Профессия
        public string ReqUserGender;                                 // Пол
        public string Department;                                    // Подразделение
        public string MeasureUnit;                                   // Ед. измерения
        public string UnionParameter;                                // Параметр для группировки (для сводной ведомости)
    }

    // Компарер по параметру для группировки
    public class UnionParameterComparer : IEqualityComparer<ClothesClaim>
    {
        public bool Equals(ClothesClaim x, ClothesClaim y)
        {
            return x.UnionParameter.Equals(y.UnionParameter);
        }

        public int GetHashCode(ClothesClaim obj)
        {
            return obj.UnionParameter.GetHashCode();
        }
    }
}
