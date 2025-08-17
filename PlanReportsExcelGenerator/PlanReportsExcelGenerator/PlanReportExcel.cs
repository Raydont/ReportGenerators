using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReportHelpers;
using Microsoft.Office.Interop.Excel;
using PlanReport.Types;

namespace Globus.PlanReportsExcel
{
    public class PlanReportExcel
    {
        public const int bookHeight = 750;
        public const int albumHeight = 495;

        // Создание верхнего блока подписей
        private int CreateSignHeader(List<SignBlock> signList, int columnsCount, bool isBook, Xls xls, ProgressForm p_form, int progressStep)
        {

            int row;
            int col;
            int cnt;
            int step;
            int rowStep = 0;
            int lastRow = 1;

            row = 1;
            col = 1;
            cnt = 0;

            signList = sortSignBlockList(signList, columnsCount, isBook);

            if (isBook)
            {
                #region Создание блоков подписей (книжного формата)
                step = 16;
                foreach (var block in signList)
                {
                    p_form.progressBarParam(progressStep);
                    string signerPost;

                    if (cnt == 3)
                    {
                        col = 1;
                        cnt = 0;
                        rowStep = rowStep + 5;
                    }

                    row = rowStep + 1;

                    if (block != null)
                    {
                        if ((block.departmentType == DepartmentType.GeneralDirector) || (block.departmentType == DepartmentType.DeputyGeneralDirector) ||
                            (block.departmentType == DepartmentType.ChiefEngineer) || (block.departmentType == DepartmentType.ChiefDesigner))
                            signerPost = block.GetDepartmentTypeChiefString() + "\r\nАО \"РКБ \"Глобус\"";
                        else
                            signerPost = block.GetDepartmentTypeChiefString();
                        
                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue(block.GetSignTypeString());
                    }
                    else
                        signerPost = String.Empty;
                    row++;

                    if (block != null)
                    {
                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue(signerPost);
                        xls[col, row].Cells.WrapText = true;
                        xls[col, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    }
                    xls[col, row].Cells.RowHeight = 45;
                    row++;
                    if (block != null)
                    {
                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue("__________" + block.fio);
                    }
                    row++;
                    if (block != null)
                    {
                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue("\"____\"____________ 20    г.");
                    }
                    row++;

                    col = col + step + 1;

                    lastRow = row;
                    cnt++;
                }
                xls[1, 1, 50, lastRow].HorizontalAlignment = XlVAlign.xlVAlignCenter;
                #endregion
            }
            else
            {
                #region Создание блоков подписей (альбомного формата)
                step = 15;
                foreach (var block in signList)
                {
                    p_form.progressBarParam(progressStep);
                    string signerPost;

                    if (cnt == 5)
                    {
                        col = 1;
                        cnt = 0;
                        rowStep = rowStep + 5;
                    }

                    row = rowStep + 1;

                    if (block != null)
                    {
                        if ((block.departmentType == DepartmentType.GeneralDirector) || (block.departmentType == DepartmentType.DeputyGeneralDirector) ||
                            (block.departmentType == DepartmentType.ChiefEngineer) || (block.departmentType == DepartmentType.ChiefDesigner))
                            signerPost = block.GetDepartmentTypeChiefString() + "\r\nАО \"РКБ \"Глобус\"";
                        else
                            signerPost = block.GetDepartmentTypeChiefString();

                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue(block.GetSignTypeString());
                    }
                    else
                        signerPost = String.Empty;

                    row++;
                    if (block != null)
                    {
                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue(signerPost);
                        xls[col, row].Cells.WrapText = true;
                        xls[col, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    }
                    xls[col, row].Cells.RowHeight = 45;
                    row++;
                    if (block != null)
                    {
                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue("__________" + block.fio);
                    }
                    row++;
                    if (block != null)
                    {
                        xls[col, row, step, 1].Merge();
                        xls[col, row].SetValue("\"____\"____________ 20    г.");
                    }
                    row++;
                    col = col + step;

                    lastRow = row;
                    cnt++;
                }
                xls[1, 1, 76, lastRow].HorizontalAlignment = XlVAlign.xlVAlignCenter;
                #endregion
            }

            if (lastRow < 11)
                lastRow = 11;
            else lastRow = lastRow + 2;

            return lastRow;
        }
        // Создание заголовка плана (графика)
        private int CreatePlanHeader(WorkHeader workHeader, bool isBook, bool newPage, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            // высота страницы формата А4
            int a4Height;
            if (isBook)
                a4Height = bookHeight;
            else
                a4Height = albumHeight;
            // Количество символов в одной строке в зависимости от формата
            int interHeight;                                        

            int row = startRow;
            int col = 1;

            if (isBook)
            {
                for (int i = 0; i < 5; i++)
                {
                    xls[1, row + i, 50, 1].Merge();
                    xls[col, row + i, 50, 1].Font.Bold = true;
                    xls[col, row + i, 50, 1].CenterText();
                }
            }
            else
            {
                for (int i = 0; i < 5; i++)
                {
                    xls[1, row + i, 76, 1].Merge();
                    xls[col, row + i, 76, 1].Font.Bold = true;
                    xls[col, row + i, 76, 1].CenterText();
                }
            }

            // Размещение названия
            xls[col, row].WrapText = true;
            xls[col, row].SetValue(workHeader.GraphName);
            // Если с новой страницы, то определяем строку, с которой начинается новая страница
            if (newPage)
            {
                interHeight = (int)xls[1, 1, 1, row].Height;
                // Проверяем, пока текущая высота не превысит размер формата А4
                while (interHeight <= a4Height)
                {
                    row++;
                    interHeight = (int)xls[1, 1, 1, row].Height;
                }
            }
            else
                row = row + 6;

            p_form.progressBarParam(progressStep);

            return row;
        }
        // Создание заголовка таблицы
        private int CreateTableHeader(WorkTypes workType, int startRow, bool isBook, Xls xls, ProgressForm p_form, int progressStep)
        {
            int row = startRow;
            int col = 1;

            p_form.progressBarParam(progressStep);

            switch (workType)
            {
                case WorkTypes.SimplePlan:
                    #region Простой план (график, мероприятия)
                    if (isBook)
                    {
                        xls[col, row, 4, 1].Merge(); xls[col, row].SetValue("№ п/п");
                        col = col + 4;
                        xls[col, row, 13, 1].Merge(); xls[col, row].SetValue("Название работы");
                        col = col + 13;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("Исполнитель");
                        col = col + 10;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("Срок");
                        col = col + 10;
                        xls[col, row, 13, 1].Merge(); xls[col, row].SetValue("Примечание");
                        xls[1, row].Cells.Rows.RowHeight = 3 * 15;
                        xls[1, row, 50, 1].Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        xls[1, row, 50, 1].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                        xls[1, row, 50, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                        xls[1, row, 50, 1].Font.Bold = true;
                    }
                    else
                    {
                        xls[col, row, 4, 1].Merge(); xls[col, row].SetValue("№ п/п");
                        col = col + 4;
                        xls[col, row, 26, 1].Merge(); xls[col, row].SetValue("Название работы");
                        col = col + 26;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("Исполнитель");
                        col = col + 10;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("Срок");
                        col = col + 10;
                        xls[col, row, 26, 1].Merge(); xls[col, row].SetValue("Примечание");
                        xls[1, row].Cells.Rows.RowHeight = 3 * 15;
                        xls[1, row, 76, 1].Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        xls[1, row, 76, 1].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                        xls[1, row, 76, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                        xls[1, row, 76, 1].Font.Bold = true;
                    }
                    row++;
#endregion
                    break;
                case WorkTypes.PreliminaryTestPlan:
                    #region График проведения предварительных испытаний
                    xls[col, row, 5, 4].Merge(); xls[col, row].SetValue("№ п/п");
                    col = col + 5;
                    xls[col, row, 18, 4].Merge(); xls[col, row].SetValue("Краткое содержание пункта программы");
                    col = col + 18;
                    xls[col, row, 10, 4].Merge(); xls[col, row].SetValue("Исполнитель\nРБ, отдел");
                    col = col + 10;
                    xls[col, row, 10, 4].Merge(); xls[col, row].SetValue("Методика разработана");
                    col = col + 10;
                    xls[col, row, 10, 4].Merge(); xls[col, row].SetValue("Испытания проведены");
                    col = col + 10;
                    xls[col, row, 10, 4].Merge(); xls[col, row].SetValue("Протокол\nсогласован\nс ВП МО и\nутвержден");
                    col = col + 10;
                    xls[col, row, 13, 4].Merge(); xls[col, row].SetValue("Примечание");
                    col = col + 13;
                    
                    xls[1, row, 76, 4].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[1, row, 76, 4].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    xls[1, row, 76, 4].Font.Bold = true;
                    xls[1, row, 76, 4].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    xls[1, row, 76, 4].WrapText = true;
                    
                    row = row + 4;
#endregion
                    break;
                case WorkTypes.DocDevelopPlan:
                    #region График разработки документации
                    xls[col, row, 5, 4].Merge(); xls[col, row].SetValue("№ п/п");
                    xls[col, row, 5, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 5;
                    xls[col, row, 10, 4].Merge(); xls[col, row].SetValue("Наименование устройства/\nдокумента"); xls[col, row].WrapText = true;
                    xls[col, row, 10, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 10;
                    xls[col, row, 3, 4].Merge(); xls[col, row].SetValue("Отдел-разработчик"); xls[col, row].Orientation = 90;
                    xls[col, row, 3, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 3;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Утвержденное ТЗ(ЗИ) выдано"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 2].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Отдел 05"); xls[col, row].WrapText = true;
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("ТЗ в смежные подразделения выдано"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 2].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Отд.-разраб."); xls[col, row].WrapText = true;
                    row = row - 3;
                    col = col + 9;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("ЭЗ разраб., ТЗ на КД выдано"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 2].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Отд.-разраб."); xls[col, row].WrapText = true;
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("СЗ об уточнении структуры в отд.906"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 2].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Отд.-разраб."); xls[col, row].WrapText = true;
                    row = row - 3;
                    col = col + 9;
                    xls[col, row, 9, 3].Merge(); xls[col, row].SetValue("ИД на ПП в РБ-128 выданы"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row = row + 3;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Отдел 20"); xls[col, row].WrapText = true;
                    row = row - 3;
                    col = col + 9;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("КД в отд.20 передана"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 2].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("РБ-218"); xls[col, row].WrapText = true;
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Комплект КД в БТД сдан"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 2].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row++;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Отдел 20"); xls[col, row].WrapText = true;
                    row = row - 3;
                    col = col + 9;
                    xls[col, row, 9, 3].Merge(); xls[col, row].SetValue("Текстовая КД (ТУ,ТУ1,И21,ТО) разработана"); xls[col, row].WrapText = true;
                    xls[col, row, 9, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row = row + 3;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Отд.-разраб."); xls[col, row].WrapText = true;
                    row = row - 3;
                    col = col + 9;
                    xls[col, row, 13, 4].Merge(); xls[col, row].SetValue("Примечание"); xls[col, row].WrapText = true;
                    xls[col, row, 13, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    
                    xls[1, row, 76, 4].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[1, row, 76, 4].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    xls[1, row, 76, 4].Font.Bold = true;
                    xls[1, row].Cells.Rows.RowHeight = 45;
                    xls[1, row + 2].Cells.Rows.RowHeight = 60;
                    row = row + 4;
#endregion
                    break;
                case WorkTypes.PeriodicTestPlan:
                    #region График проведения периодических испытаний
                    xls[col, row, 4, 5].Merge(); xls[col, row].SetValue("№ п/п");
                    xls[col, row, 4, 5].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 4;
                    xls[col, row, 17, 5].Merge(); xls[col, row].SetValue("Наименование устройства\nи шифр");
                    xls[col, row, 17, 5].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 17;
                    xls[col, row, 10, 5].Merge(); xls[col, row].SetValue("Акт ПИ");
                    xls[col, row, 10, 5].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 10;
                    xls[col, row, 10, 5].Merge(); xls[col, row].SetValue("Дата окончания\nдействия акта ПИ");
                    xls[col, row, 10, 5].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 10;
                    xls[col, row, 25, 1].Merge(); xls[col, row].SetValue("Изготовление устройства");
                    xls[col, row, 25, 1].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row++;
                    xls[col, row, 15, 4].Merge(); xls[col, row].SetValue("График изготовления");
                    xls[col, row, 15, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    col = col + 15;
                    xls[col, row, 10, 4].Merge(); xls[col, row].SetValue("Срок изготовления и сдачи ц.75, отд.20"); xls[col, row].WrapText = true;
                    xls[col, row, 10, 4].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    row = row - 1;
                    col = col + 10;
                    xls[col, row, 10, 5].Merge(); xls[col, row].SetValue("Срок проведения ПИ, отд.20"); xls[col, row].WrapText = true;
                    xls[col, row, 10, 5].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                    
/*                    xls[1, row, 76, 5].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
                    xls[1, row, 76, 5].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;*/
                    xls[1, row, 76, 5].CenterText();
                    xls[1, row, 76, 5].Font.Bold = true;
                    
                    row = row + 5;
#endregion
                    break;
                case WorkTypes.ProductionPlan:
                    #region График изготовления
                    xls[col, row, 5, 2].Merge(); xls[col, row].SetValue("№ п/п"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 5;
                    xls[col, row, 7, 2].Merge(); xls[col, row].SetValue("Наименование и шифр"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 7;
                    xls[col, row, 5, 2].Merge(); xls[col, row].SetValue("Децимальный номер"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 5;
                    xls[col, row, 2, 2].Merge(); xls[col, row].SetValue("Кол-во"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 2;
                    xls[col, row, 3, 2].Merge(); xls[col, row].SetValue("Разработчик"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 6, 1].Merge(); xls[col, row].SetValue("Н.-з. в б.219\r\n(разраб.)"); xls[col, row].WrapText = true;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    row++;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("Дата"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("Номер"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    row--;
                    xls[col, row, 4, 2].Merge(); xls[col, row].SetValue("Н.-з. и к.д. в отд.60\r\n(б.219)"); xls[col, row].WrapText = true;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                    col = col + 4;
                    xls[col, row, 3, 2].Merge(); xls[col, row].SetValue("Технол. обр-ка н.-з. (отд.60)"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 3, 2].Merge(); xls[col, row].SetValue("Нормир-е (бюро 904)"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 6, 1].Merge(); xls[col, row].SetValue("Выдача"); 
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    row++;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("Материалы (отд.914)"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("КРЭ (отд.914)"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    row--;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("Изготовление"); 
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    row++;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("800"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("850"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("750"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    row--;
                    xls[col, row, 3, 2].Merge(); xls[col, row].SetValue("Настройка / Проверка"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 3, 2].Merge(); xls[col, row].SetValue("Исполнитель настр."); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 3, 2].Merge(); xls[col, row].SetValue("Дефектировка"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 4, 2].Merge(); xls[col, row].SetValue("Сдача"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 4;
                    xls[col, row, 3, 2].Merge(); xls[col, row].SetValue("Исполнитель сдачи"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    col = col + 3;
                    xls[col, row, 7, 2].Merge(); xls[col, row].SetValue("Примечание"); xls[col, row].Orientation = 90;
                    xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;

                    xls[col, row].Cells.RowHeight = 45;
                    xls[col, row + 1].Cells.RowHeight = 120;
                    xls[1, row, 76, 2].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    xls[1, row, 76, 2].Font.Bold = true;
                    xls[1, row, 76, 2].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

                    row = row + 2;
                    #endregion
                    break;
            }

            return row;
        }
        // Создание цифрового блока таблицы
        private int CreateDigitalHeader(WorkTypes workType, int startRow, bool isBook, Xls xls)
        {
            int row = startRow;
            int col = 1;

            if (isBook)
            {
                xls[1, startRow, 50, 1].Font.Bold = true;
                xls[1, startRow, 50, 1].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                xls[1, startRow, 50, 1].NumberFormat = "@";
                xls[1, startRow, 50, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }
            else
            {
                xls[1, startRow, 76, 1].Font.Bold = true;
                xls[1, startRow, 76, 1].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                xls[1, startRow, 76, 1].NumberFormat = "@";
                xls[1, startRow, 76, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }

            switch(workType)
            {
                case WorkTypes.SimplePlan:
                    #region Простой план (график, мероприятия)
                    if (isBook)
                    {
                        xls[col, row, 4, 1].Merge(); xls[col, row].SetValue("1");
                        col = col + 4;
                        xls[col, row, 13, 1].Merge(); xls[col, row].SetValue("2");
                        col = col + 13;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("3");
                        col = col + 10;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("4");
                        col = col + 10;
                        xls[col, row, 13, 1].Merge(); xls[col, row].SetValue("5");
                    }
                    else
                    {
                        xls[col, row, 4, 1].Merge(); xls[col, row].SetValue("1");
                        col = col + 4;
                        xls[col, row, 26, 1].Merge(); xls[col, row].SetValue("2");
                        col = col + 26;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("3");
                        col = col + 10;
                        xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("4");
                        col = col + 10;
                        xls[col, row, 26, 1].Merge(); xls[col, row].SetValue("5");
                    }
#endregion
                    break;
                case WorkTypes.PreliminaryTestPlan:
                    #region График проведения предварительных испытаний
                    xls[col, row, 5, 1].Merge(); xls[col, row].SetValue("1");
                    col = col + 5;                    
                    xls[col, row, 18, 1].Merge(); xls[col, row].SetValue("2");
                    col = col + 18;                    
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("3");
                    col = col + 10;                    
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("4");
                    col = col + 10;                    
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("5");
                    col = col + 10;                    
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("6");
                    col = col + 10;                    
                    xls[col, row, 13, 1].Merge(); xls[col, row].SetValue("7");
#endregion
                    break;
                case WorkTypes.DocDevelopPlan:
                    #region График разработки документации
                    xls[col, row, 5, 1].Merge(); xls[col, row].SetValue("1");
                    col = col + 5;
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("2");
                    col = col + 10;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("3");
                    col = col + 3;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("4 / 5");
                    col = col + 9;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("6 / 7");
                    col = col + 9;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("8");
                    col = col + 9;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("9 / 10");
                    col = col + 9;
                    xls[col, row, 9, 1].Merge(); xls[col, row].SetValue("11");
                    col = col + 9;
                    xls[col, row, 13, 1].Merge(); xls[col, row].SetValue("12");
#endregion
                    break;
                case WorkTypes.PeriodicTestPlan:
                    #region График проведения периодических испытаний
                    xls[col, row, 4, 1].Merge(); xls[col, row].SetValue("1");
                    col = col + 4;
                    xls[col, row, 17, 1].Merge(); xls[col, row].SetValue("2");
                    col = col + 17;
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("3");
                    col = col + 10;
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("4");
                    col = col + 10;
                    xls[col, row, 15, 1].Merge(); xls[col, row].SetValue("5");
                    col = col + 15;
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("6");
                    col = col + 10;
                    xls[col, row, 10, 1].Merge(); xls[col, row].SetValue("7");
                    #endregion
                    break;
                case WorkTypes.ProductionPlan:
                    #region График изготовления
                    xls[col, row, 5, 1].Merge(); xls[col, row].SetValue("1");
                    col = col + 5;
                    xls[col, row, 7, 1].Merge(); xls[col, row].SetValue("2");
                    col = col + 7;
                    xls[col, row, 5, 1].Merge();
                    col = col + 5;
                    xls[col, row, 2, 1].Merge(); xls[col, row].SetValue("3");
                    col = col + 2;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("4");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("5");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("6");
                    col = col + 3;
                    xls[col, row, 4, 1].Merge(); xls[col, row].SetValue("7");
                    col = col + 4;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("8");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("9");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("10");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("11");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("12");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("13");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("14");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("15");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("16");
                    col = col + 3;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("17");
                    col = col + 3;
                    xls[col, row, 4, 1].Merge(); xls[col, row].SetValue("18");
                    col = col + 4;
                    xls[col, row, 3, 1].Merge(); xls[col, row].SetValue("19");
                    col = col + 3;
                    xls[col, row, 7, 1].Merge(); xls[col, row].SetValue("20");
#endregion
                    break;
            }

            row++;

            return row;
        }
        // Создание нижнего блока подписей
        private void CreateSignFooter(List<SignBlock> signFooterList, bool isBook, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            int blockHeight = (signFooterList.Count * 2 - 1) * 15;
            int row = BlockStartPosition(startRow, blockHeight, isBook, xls);
            int col = 7;

            foreach (var block in signFooterList)
            {
                p_form.progressBarParam(progressStep);
                xls[col, row].SetValue(block.GetDepartmentTypeChiefString());
                xls[35, row].SetValue(block.fio);

                row = row + 2;
            }
        }
        // Создание перечня работ
        private int CreatePlanBody(List<WorkParameters> list, bool isBook, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            int symbolsCount;                                                       // Количество символов в одной строке в зависимости от формата
            if (isBook)
                symbolsCount = 60;
            else
                symbolsCount = 100;
            int lastRow = 0;

            int heightCoeff;
            int cellHeight;

            int row = startRow;
            int col = 3;

            foreach (var block in list)
            {
                p_form.progressBarParam(progressStep);

                // Высота строки с наименованием работы
                heightCoeff = GetHeightCoeff(block.WorkSummary.Length, symbolsCount);
                cellHeight = 15 * heightCoeff;

                row = BlockStartPosition(row, cellHeight + 15, isBook, xls);

                #region Заполнение строки плана
                // Пункт плана
                xls[col, row, 4, 1].Merge();
                xls[col, row].NumberFormat = "@";
                if (block.IsHeader)
                    xls[col, row].Font.Bold = true;
                xls[col, row].SetValue(block.PlanItem);
                col = col + 4;
                // Название работы
                if (isBook)
                    xls[col, row, 43, 1].Merge();
                else
                    xls[col, row, 69, 1].Merge();
                if (block.IsHeader)
                    xls[col, row].Font.Bold = true;
                xls[col, row].SetValue(block.WorkSummary);
                xls[col, row].WrapText = true;
                xls[col, row].Cells.RowHeight = cellHeight;

                // Исполнитель и срок исполнения
                if (!block.IsHeader)
                {
                    row++; 
                    col = 7;
                    // Высота строки со списком исполнителей
                    heightCoeff = block.Executor.Count(t => t == '\n');
                    cellHeight = 15 * heightCoeff;
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue("Исп.:");
                    xls[col + 3, row, 20, 1].Merge();
                    xls[col + 3, row].SetValue(block.Executor);
                    xls[col + 3, row].WrapText = true;
                    xls[col, row].Cells.RowHeight = cellHeight;
                    if (isBook)
                        col = col + 35;
                    else
                        col = col + 50;
                    xls[col, row].NumberFormat = "@";
                    xls[col, row].SetValue("Срок: " + block.Deadline);
                }

                #endregion

                lastRow = row;
                row = row + 2;
                col = 3;
            }

            if (isBook)
                xls[3, startRow, 47, lastRow - startRow + 1].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
            else
                xls[3, startRow, 73, lastRow - startRow + 1].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;

            lastRow = lastRow + 2;

            return lastRow;
        }
        // Заполнение простой табличной формы
        private int CreateTableBody(List<WorkParameters> list, WorkTypes workType, bool isBook, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            // Количество символов в одной строке в зависимости от формата
            int symbolsCount;
            if (isBook)
                symbolsCount = 21;
            else
                symbolsCount = 42;
            int heightCoeff;
            int cellHeight;

            int row = startRow;
            int col = 1;

            foreach (var block in list)
            {
                p_form.progressBarParam(progressStep);
                #region Установка высоты строки таблицы
                List<int> rowHeights = new List<int>();
                // Стандартная высота строки
                rowHeights.Add(15);
                if (!block.IsHeader)
                {
                    heightCoeff = GetHeightCoeff(block.WorkSummary.Length, symbolsCount);
                    // Предполагаемый размер строки с названием работы
                    cellHeight = 15 * (heightCoeff + 1);
                    rowHeights.Add(cellHeight);
                    if (block.Comment != null && block.Comment != String.Empty)
                    {
                        heightCoeff = GetHeightCoeff(block.Comment.Length, symbolsCount);
                        // Предполагаемый размер строки с примечанием
                        cellHeight = 15 * (heightCoeff + 1);
                        rowHeights.Add(cellHeight);
                    }
                    // Высота строки со списком исполнителей
                    heightCoeff = block.Executor.Count(t => t == '\n');
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }
                // Высота строки - максимальная из вышеперечисленных
                cellHeight = rowHeights.Max();

                row = LastRowHeight(row, cellHeight, workType, isBook, xls);

                if (cellHeight > 15)
                    xls[col, row].Cells.RowHeight = cellHeight;
                #endregion

                if (block.IsHeader)
                {
                    if (isBook)
                        xls[1, row, 50, 1].Merge();
                    else
                        xls[1, row, 76, 1].Merge();
                    xls[1, row].Font.Bold = true;
                    xls[1, row].CenterText();
                    xls[1, row].SetValue(block.WorkSummary);
                }
                else
                {
                    #region Заполнение строки таблицы
                    if (isBook)
                        xls[1, row, 50, 1].NumberFormat = "@";
                    else
                        xls[1, row, 76, 1].NumberFormat = "@";
                    // Пункт плана
                    xls[col, row, 4, 1].Merge();
                    xls[col, row].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    xls[col, row].SetValue(block.PlanItem);
                    col = col + 4;

                    // Название работы
                    if (isBook)
                        xls[col, row, 13, 1].Merge();
                    else
                        xls[col, row, 26, 1].Merge();
                    xls[col, row].SetValue(block.WorkSummary);
                    xls[col, row].WrapText = true;
                    if (isBook)
                        col = col + 13;
                    else
                        col = col + 26;

                    // Исполнитель
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    xls[col, row].SetValue(block.Executor);
                    col = col + 10;

                    // Срок исполнения
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    xls[col, row].SetValue(block.Deadline);
                    col = col + 10;

                    // Примечание
                    if (isBook)
                        xls[col, row, 13, 1].Merge();
                    else
                        xls[col, row, 26, 1].Merge();
                    xls[col, row].SetValue(block.Comment);
                    xls[col, row].WrapText = true;
                    #endregion
                }

                col = 1;
                row++;
            }

            if (isBook)
            {
                xls[1, startRow, 50, row - startRow].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
                xls[1, startRow, 50, row - startRow].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                xls[1, startRow - 1, 50, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                xls[1, startRow, 50, row - startRow].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }
            else
            {
                xls[1, startRow, 76, row - startRow].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
                xls[1, startRow, 76, row - startRow].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
                xls[1, startRow - 1, 76, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
                xls[1, startRow, 76, row - startRow].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            }

            row = row + 2;
            
            return row;
        }
        // Заполнение табличной формы графика разработки КД
        private int CreateTableBody(List<WorkDocDevelopParameters> list, WorkTypes workType, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            int row = startRow;
            int col = 1;
            int symbolsCount;
            int cellHeight;
            int heightCoeff;

            foreach (var block in list)
            {
                p_form.progressBarParam(progressStep);

                #region Установка высоты строки таблицы
                List<int> rowHeights = new List<int>();
                // Стандартная высота строки
                rowHeights.Add(15);
                if (!block.IsHeader)
                {
                    // Высота ячейки с пунктом структуры
                    symbolsCount = 9;
                    heightCoeff = GetHeightCoeff(block.StructItem.Length, symbolsCount);
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                    // Высота ячейки с наименованием
                    symbolsCount = 15;
                    heightCoeff = GetHeightCoeff(block.DeviceName.Length, symbolsCount);
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                    // Высота строки со списком исполнителей
                    heightCoeff = block.Developer.Count(t => t == '\n');
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }
                // Высота ячейки с комментариями
                if (block.Comments != null && block.Comments != String.Empty)
                {
                    symbolsCount = 20;
                    heightCoeff = GetHeightCoeff(block.Comments.Length, symbolsCount);
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }

                // Максимальная высота строки
                cellHeight = rowHeights.Max();

                // Установка высоты строки таблицы
                row = LastRowHeight(row, cellHeight, workType, false, xls);
                if (cellHeight > 15)
                    xls[col, row + 1].Cells.Rows.RowHeight = cellHeight - 15;
                #endregion

                xls[1, row, 76, 1].NumberFormat = "@";

                if (block.IsHeader)
                {
                    xls[1, row, 76, 1].Merge();
                    xls[1, row].CenterText();
                    xls[1, row].Font.Bold = true;
                    xls[1, row].SetValue(block.DeviceName);
                    row = row + 1;
                }
                else
                {
                    #region Заполнение строки таблицы
                    // Пункт структуры
                    xls[col, row, 5, 2].Merge();
                    xls[col, row].SetValue(block.StructItem);
                    col = col + 5;
                    // Наименование устройства / документа
                    xls[col, row, 10, 2].Merge();
                    xls[col, row].SetValue(block.DeviceName);
                    xls[col, row].WrapText = true;
                    col = col + 10;
                    // Отдел - разработчик
                    xls[col, row, 3, 2].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.Developer);
                    col = col + 3;
                    // Утвержденное ТЗ выдано
                    xls[col, row, 9, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.TechTaskConfirmDate);
                    row++;
                    // ТЗ в смежные подразделения
                    xls[col, row, 9, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.TechTaskStatementDate);
                    row--;
                    col = col + 9;
                    // Э3 разработано, КД выдано
                    xls[col, row, 9, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.StructDevelopDate);
                    row++;
                    // СЗ об уточнении структуры
                    xls[col, row, 9, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.ServiceDocTo906Date);
                    row--;
                    col = col + 9;
                    // ИД на ПП выданы
                    xls[col, row, 9, 2].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.InitialDataDate);
                    col = col + 9;
                    // КД в отд.20 выдана
                    xls[col, row, 9, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.EngineeringDocDate);
                    row++;
                    // Комплект КД в БТД сдан
                    xls[col, row, 9, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.EngineeringDocComplectDate);
                    row--;
                    col = col + 9;
                    // Текстовая КД разработана
                    xls[col, row, 9, 2].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.TextEngineeringDocDate);
                    col = col + 9;
                    // Примечания
                    xls[col, row, 13, 2].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.Comments);
                    #endregion
                    row = row + 2;
                }
                col = 1;
            }

            xls[1, startRow, 76, row - startRow].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, startRow, 76, row - startRow].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, startRow, 76, row - startRow].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            xls[15, startRow, 48, row - startRow].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;

            return row;
        }
        // Заполнение табличной формы графика периодических испытаний
        private int CreateTableBody(List<WorkPeriodicTestParameters> list, WorkTypes workType, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            int row = startRow;
            int col = 1;
            int symbolsCount;
            int heightCoeff;
            int cellHeight;

            foreach (var block in list)
            {
                p_form.progressBarParam(progressStep);

                #region Установка высоты строки таблицы
                List<int> rowHeights = new List<int>();
                // Стандартная высота строки
                rowHeights.Add(15);
                // Высота строки с наименованием устройства
                if (!block.IsHeader)
                {
                    symbolsCount = 25;
                    heightCoeff = GetHeightCoeff(block.DeviceName.Length, symbolsCount);
                    cellHeight = 15 * (heightCoeff + 1);
                    rowHeights.Add(cellHeight);
                }
                // Максимальная высота строки
                cellHeight = rowHeights.Max();

                row = LastRowHeight(row, cellHeight, workType, false, xls);

                // Установка высоты строки таблицы
                if (cellHeight > 15)
                    xls[col, row].Cells.Rows.RowHeight = cellHeight;
                #endregion

                xls[1, row, 76, 1].NumberFormat = "@";

                if (block.IsHeader)
                {
                    xls[1, row, 76, 1].Merge();
                    xls[1, row].Font.Bold = true;
                    xls[1, row].SetValue(block.DeviceName);
                    xls[1, row].CenterText();
                }
                else
                {
                    #region Заполнение строки таблицы
                    // Пункт плана
                    xls[col, row, 4, 1].Merge();
                    xls[col, row].SetValue(block.StructItem);
                    col = col + 4;
                    // Наименование устройства
                    xls[col, row, 17, 1].Merge();
                    xls[col, row].SetValue(block.DeviceName);
                    col = col + 17;
                    // Акт ПИ
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.ActPT);
                    col = col + 10;
                    // Дата окончания действия акта ПИ
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.ActPTFinishDate);
                    col = col + 10;
                    // График изготовления
                    xls[col, row, 15, 1].Merge();
                    xls[col, row].SetValue(block.ProductionGraph);
                    col = col + 15;
                    // Срок изготовления
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.ProductionDate);
                    col = col + 10;
                    // Срок проведения ПИ
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.PTDate);
                    #endregion
                }

                col = 1;
                row++;
            }

            xls[1, startRow, 76, row - startRow].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, startRow, 76, row - startRow].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, startRow, 76, row - startRow].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            xls[1, startRow, 76, row - startRow].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;

            row = row + 2;
            return row;
        }
        // Заполнение табличной формы графика предварительных испытаний
        private int CreateTableBody(List<WorkPreliminaryTestParameters> list, WorkTypes workType, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            int row = startRow;
            int col = 1;
            int symbolsCount;
            int heightCoeff;
            int cellHeight;

            foreach (var block in list)
            {
                p_form.progressBarParam(progressStep);

                #region Установка высоты строки таблицы
                List<int> rowHeights = new List<int>();
                // Стандартная высота строки
                rowHeights.Add(15);
                // Высота строки с кратким содержанием
                if (!block.IsHeader)
                {
                    symbolsCount = 30;
                    heightCoeff = GetHeightCoeff(block.PlanItemName.Length, symbolsCount);
                    cellHeight = 15 * (heightCoeff + 1);
                    rowHeights.Add(cellHeight);
                    // Высота строки со списком исполнителей
                    heightCoeff = block.Executor.Count(t => t == '\n');
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }
                // Высота строки с Примечанием
                if (block.Comments != null && block.Comments != String.Empty)
                {
                    symbolsCount = 25;
                    heightCoeff = GetHeightCoeff(block.Comments.Length, symbolsCount);
                    cellHeight = 15 * (heightCoeff + 1);
                    rowHeights.Add(cellHeight);
                }

                // Максимальная высота строки
                cellHeight = rowHeights.Max();

                row = LastRowHeight(row, cellHeight, workType, false, xls);
                // Установка высоты строки таблицы
                if (cellHeight > 15)
                    xls[col, row].Cells.Rows.RowHeight = cellHeight;
                #endregion

                xls[1, row, 76, 1].NumberFormat = "@";

                if (block.IsHeader)
                {
                    xls[1, row, 76, 1].Merge();
                    xls[1, row].Font.Bold = true;
                    xls[1, row].SetValue(block.PlanItemName);
                    xls[1, row].CenterText();
                }
                else
                {
                    #region Заполнение строки таблицы
                    // Пункт программы
                    xls[col, row, 5, 1].Merge();
                    xls[col, row].SetValue(block.PlanItem);
                    col = col + 5;
                    // Краткое содержание пункта программы
                    xls[col, row, 18, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.PlanItemName);
                    col = col + 18;
                    // Исполнитель
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.Executor);
                    col = col + 10;
                    // Методика разработана
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.MethodologyDate);
                    col = col + 10;
                    // Испытания проведены
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.TestDate);
                    col = col + 10;
                    // Протокол согласован
                    xls[col, row, 10, 1].Merge();
                    xls[col, row].SetValue(block.ProtocolConfirmDate);
                    col = col + 10;
                    // Примечания
                    xls[col, row, 13, 1].Merge();
                    xls[col, row].SetValue(block.Comments);
                    #endregion
                }
                col = 1;
                row++;
            }

            xls[1, startRow, 76, row - startRow].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, startRow, 76, row - startRow].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, startRow, 76, row - startRow].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, startRow, 76, row - startRow].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

            return row;
        }
        // Заполнение табличной формы графика изготовления
        private int CreateTableBody(List<WorkProductionParameters> list, WorkTypes workType, int startRow, Xls xls, ProgressForm p_form, int progressStep)
        {
            int row = startRow;
            int col = 1;
            int symbolsCount;
            int heightCoeff;
            int cellHeight;

            foreach (var block in list)
            {
                p_form.progressBarParam(progressStep);

                #region Установка высоты строки таблицы
                List<int> rowHeights = new List<int>();
                // Стандартная высота строки
                rowHeights.Add(15);
                // Высота строки с кратким содержанием
                if (!block.IsHeader)
                {
                    symbolsCount = 11;
                    heightCoeff = GetHeightCoeff(block.Name.Length, symbolsCount);
                    cellHeight = 15 * (heightCoeff + 1);
                    rowHeights.Add(cellHeight);
                }
                // Высота строки с Примечанием
                if ((block.Comment != null) && (block.Comment != String.Empty))
                {
                    symbolsCount = 11;
                    heightCoeff = GetHeightCoeff(block.Comment.Length, symbolsCount);
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }
                // Высота строки с Исполнителями настройки
                if ((block.SettingExecutor != null) && (block.SettingExecutor != String.Empty))
                {
                    heightCoeff = block.SettingExecutor.Count(t => t == '\n');
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }
                // Высота строки с Исполнителями сдачи
                if ((block.DeliveryExecutor != null) && (block.DeliveryExecutor != String.Empty))
                {
                    heightCoeff = block.DeliveryExecutor.Count(t => t == '\n');
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }
                // Высота строки с Децимальным номером
                if ((block.Denotation != null) && (block.Denotation != String.Empty))
                {
                    heightCoeff = block.Denotation.Count(t => t == '\n');
                    cellHeight = 15 * heightCoeff;
                    rowHeights.Add(cellHeight);
                }

                // Максимальная высота строки
                cellHeight = rowHeights.Max();

                row = LastRowHeight(row, cellHeight, workType, false, xls);
                // Установка высоты строки таблицы
                if (cellHeight > 15)
                    xls[col, row].Cells.Rows.RowHeight = cellHeight;
                #endregion

                xls[1, row, 76, 1].NumberFormat = "@";

                if (block.IsHeader)
                {
                    xls[1, row, 76, 1].Merge();
                    xls[1, row, 76, 1].Font.Size = 10;
                    xls[1, row].Font.Bold = true;
                    xls[1, row].SetValue(block.Name);
                    xls[1, row].CenterText();
                }
                else
                {
                    #region Заполнение строки таблицы
                    xls[1, row, 76, 1].Font.Size = 9;
                    // Пункт программы
                    xls[col, row, 5, 1].Merge();
                    xls[col, row].SetValue(block.PlanItem);
                    col = col + 5;
                    // Наименование и шифр
                    xls[col, row, 7, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.Name);
                    col = col + 7;
                    // Децимальный номер
                    xls[col, row, 5, 1].Merge();
                    xls[col, row].SetValue(block.Denotation);
                    col = col + 5;
                    // Количество
                    xls[col, row, 2, 1].Merge();
                    xls[col, row].SetValue(block.Amount);
                    col = col + 2;
                    // Разработчик
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.Executor);
                    col = col + 3;
                    // Дата
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.StartDate);
                    col = col + 3;
                    // Номер
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.DocNumber);
                    col = col + 3;
                    // Н.з. в отд.60
                    xls[col, row, 4, 1].Merge();
                    xls[col, row].SetValue(block.Bureau219Date);
                    col = col + 4;
                    // Технологич.обр-ка н.-з.
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.Dep60Date);
                    col = col + 3;
                    // Нормирование
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.Bureau904Date);
                    col = col + 3;
                    // Материалы
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.MaterialsDate);
                    col = col + 3;
                    // КРЭ
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.OthersDate);
                    col = col + 3;
                    // 800
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.Sector800Date);
                    col = col + 3;
                    // 850
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.Sector850Date);
                    col = col + 3;
                    // 750
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.Sector750Date);
                    col = col + 3;
                    // Настройка / Проверка
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.SettingDate);
                    col = col + 3;
                    // Исполнитель проверки
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.SettingExecutor);
                    col = col + 3;
                    // Дефектировка
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.DefectiveDate);
                    col = col + 3;
                    // Сдача
                    xls[col, row, 4, 1].Merge();
                    xls[col, row].SetValue(block.DeliveryDate);
                    col = col + 4;
                    // Исполнитель сдачи
                    xls[col, row, 3, 1].Merge();
                    xls[col, row].SetValue(block.DeliveryExecutor);
                    col = col + 3;
                    // Примечание
                    xls[col, row, 7, 1].Merge();
                    xls[col, row].WrapText = true;
                    xls[col, row].SetValue(block.Comment);
                    #endregion
                }
                col = 1;
                row++;
            }

            xls[1, startRow, 76, row - startRow].Cells.VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[18, startRow, 51, row - startRow].Cells.HorizontalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, startRow, 76, row - startRow].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[1, startRow, 76, row - startRow].Border(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);

            return row;
        }


        // Создание простого перечня работ
        public void CreatePlanReport(List<SignBlock> signBlockList, int columnsCount, List<SignBlock> footerSignList, WorkHeader workheader, WorkTypes workType, List<WorkParameters> workList, bool isBook, Xls xls, ProgressForm p_form, int progressStep)
        {
            PageSetup(xls, isBook);
            int lastRow = CreateSignHeader(signBlockList, columnsCount, isBook, xls, p_form, progressStep);
            lastRow = CreatePlanHeader(workheader, isBook, false, lastRow, xls, p_form, progressStep);
            lastRow = CreatePlanBody(workList, isBook, lastRow, xls, p_form, progressStep);
            CreateSignFooter(footerSignList, isBook, lastRow + 2, xls, p_form, progressStep);
        }
        // Создание простой табличной формы
        public void CreatePlanReport(List<SignBlock> signBlockList, int columnsCount,List<SignBlock> footerSignList, WorkHeader workheader, WorkTypes workType, List<WorkParameters> workList, bool isBook, bool newPage, Xls xls, ProgressForm p_form, int progressStep)
        {
            PageSetup(xls, isBook);
            int lastRow = CreateSignHeader(signBlockList, columnsCount, isBook, xls, p_form, progressStep);
            lastRow = CreatePlanHeader(workheader, isBook, newPage, lastRow, xls, p_form, progressStep);
            lastRow = CreateTableHeader(workType, lastRow, isBook, xls, p_form, progressStep);
            lastRow = CreateDigitalHeader(workType, lastRow, isBook, xls);
            lastRow = CreateTableBody(workList, workType, isBook, lastRow, xls, p_form, progressStep);
            CreateSignFooter(footerSignList, isBook, lastRow + 2, xls, p_form, progressStep);
        }
        // Создание табличной формы для графика разработки КД
        public void CreatePlanReport(List<SignBlock> signBlockList, int columnsCount, List<SignBlock> footerSignList, WorkHeader workheader, WorkTypes workType, List<WorkDocDevelopParameters> workList, Xls xls, ProgressForm p_form, int progressStep)
        {
            PageSetup(xls, false);
            int lastRow = CreateSignHeader(signBlockList, columnsCount, false, xls, p_form, progressStep);
            lastRow = CreatePlanHeader(workheader, false, true, lastRow, xls, p_form, progressStep);
            lastRow = CreateTableHeader(workType, lastRow, false, xls, p_form, progressStep);
            lastRow = CreateDigitalHeader(workType, lastRow, false, xls);
            lastRow = CreateTableBody(workList, workType, lastRow, xls, p_form, progressStep);
            CreateSignFooter(footerSignList, false, lastRow + 2, xls, p_form, progressStep);
        }
        // Создание табличной формы для графика проведения периодических испытаний
        public void CreatePlanReport(List<SignBlock> signBlockList, int columnsCount, List<SignBlock> footerSignList, WorkHeader workheader, WorkTypes workType, List<WorkPeriodicTestParameters> workList, Xls xls, ProgressForm p_form, int progressStep)
        {
            PageSetup(xls, false);
            int lastRow = CreateSignHeader(signBlockList, columnsCount, false, xls, p_form, progressStep);
            lastRow = CreatePlanHeader(workheader, false, true, lastRow, xls, p_form, progressStep);
            lastRow = CreateTableHeader(workType, lastRow, false, xls, p_form, progressStep);
            lastRow = CreateDigitalHeader(workType, lastRow, false, xls);
            lastRow = CreateTableBody(workList, workType, lastRow, xls, p_form, progressStep);
            CreateSignFooter(footerSignList, false, lastRow + 2, xls, p_form, progressStep);
        }
        // Создание табличной формы для графика проведения предварительных испытаний
        public void CreatePlanReport(List<SignBlock> signBlockList, int columnsCount, List<SignBlock> footerSignList, WorkHeader workheader, WorkTypes workType, List<WorkPreliminaryTestParameters> workList, Xls xls, ProgressForm p_form, int progressStep)
        {
            PageSetup(xls, false);
            int lastRow = CreateSignHeader(signBlockList, columnsCount, false, xls, p_form, progressStep);
            lastRow = CreatePlanHeader(workheader, false, true, lastRow, xls, p_form, progressStep);
            lastRow = CreateTableHeader(workType, lastRow, false, xls, p_form, progressStep);
            lastRow = CreateDigitalHeader(workType, lastRow, false, xls);
            lastRow = CreateTableBody(workList, workType, lastRow, xls, p_form, progressStep);
            CreateSignFooter(footerSignList, false, lastRow + 2, xls, p_form, progressStep);
        }
        // Создание табличной формы для графика изготовления
        public void CreatePlanReport(List<SignBlock> signBlockList, int columnsCount, List<SignBlock> footerSignList, WorkHeader workheader, WorkTypes workType, List<WorkProductionParameters> workList, Xls xls, ProgressForm p_form, int progressStep)
        {
            PageSetup(xls, false);
            int lastRow = CreateSignHeader(signBlockList, columnsCount, false, xls, p_form, progressStep);
            lastRow = CreatePlanHeader(workheader, false, true, lastRow, xls, p_form, progressStep);
            lastRow = CreateTableHeader(workType, lastRow, false, xls, p_form, progressStep);
            lastRow = CreateDigitalHeader(workType, lastRow, false, xls);
            lastRow = CreateTableBody(workList, workType, lastRow, xls, p_form, progressStep);
            CreateSignFooter(footerSignList, false, lastRow + 2, xls, p_form, progressStep);
        }


        // Вычисление стартовой позиции блока (для перечня работ и нижнего блока подписей)
        private int BlockStartPosition(int startRow, int blockHeight, bool isBook, Xls xls)
        {
            int row = 0;
            // высота страницы формата А4
            int a4Height;

            if (isBook)
                a4Height = bookHeight;
            else
                a4Height = albumHeight;

            // высота тела документа без текущей записи
            int currentHeight = (int)xls[1, 1, 1, startRow - 1].Height;
            // сколько листов занимает тело документа без текущей записи
            int currentPages = GetHeightCoeff(currentHeight, a4Height);

            // высота тела документа с текущей записью
            int nextHeight = currentHeight + blockHeight;
            // сколько листов занимает тело документа с текущей записью
            int nextPages = GetHeightCoeff(nextHeight, a4Height);

            // Если документ с блоком подписей занимает больше листов, чем тело документа,
            // то размещаем блок подписей на другой странице
            if (nextPages > currentPages)
            {
                row = startRow;
                // Текущая высота тела документа
                int interHeight = (int)xls[1, 1, 1, row].Height;
                // Проверяем, пока текущая высота не превысит размер формата А4
                while (interHeight <= a4Height * currentPages)
                {
                    row++;
                    interHeight = (int)xls[1, 1, 1, row].Height;
                }
            }
            else
                row = startRow;

            return row;
        }
        // Задание ширины последней на листе строки
        private int LastRowHeight(int lastRow, int cellHeight, WorkTypes workType, bool isBook, Xls xls)
        {
            int row = lastRow;

            // высота страницы формата А4
            int a4Height;

            if (isBook)
                a4Height = bookHeight;
            else
                a4Height = albumHeight;

            // высота тела документа без текущей записи
            int currentHeight = (int)xls[1, 1, 1, lastRow - 1].Height;
            // сколько листов занимает тело документа без текущей записи
            int currentPages = GetHeightCoeff(currentHeight, a4Height);

            // высота тела документа с текущей записью
            int nextHeight = currentHeight + cellHeight;
            // сколько листов занимает тело документа с текущей записью
            int nextPages = GetHeightCoeff(nextHeight, a4Height);

            // Если документ с текущей записью занимает больше листов, чем документ без текущей записи,
            // то высота предпоследней строки увеличивается до конца страницы
            if (nextPages > currentPages)
            {
                // Остаток до конца страницы
                int restHeight = (a4Height * currentPages) - currentHeight;
                // Текущая высота предпоследней строки
                int preLastRowHeight = (int)xls[1, lastRow - 1].Cells.Rows.Height;
                // Новая высота предпоследнй строки
                xls[1, lastRow - 1].Cells.RowHeight = preLastRowHeight + restHeight;
                CreateDigitalHeader(workType, lastRow, isBook, xls);
                row++;
            }

            return row;
        }
        // Вычисление количества листов, на которых располагается блок
        private int GetHeightCoeff(int length, int symbolsCount)
        {
            int coeff = 0;

            int ratio = Math.Abs(length / symbolsCount);
            double preciseRatio = (double)(length / (double)symbolsCount);

            if ((preciseRatio - ratio) < 0.000001)
                coeff = ratio;
            else
                coeff = ratio + 1;

            return coeff;
        }
        // Пересортировка списка подписей в зависимости от позиции, количества колонок и ориентации страницы
        private List<SignBlock> sortSignBlockList(List<SignBlock> list, int columnCount, bool isBook)
        {
            List<SignBlock> sortList = new List<SignBlock>();

            list = list.OrderBy(t => t.position).ToList();
            int lastPosition = list[list.Count - 1].position;
            int rowCount;
            // Параметр "Позиция" элемента списка
            int position;
            // Индекс в новом списке
            int index;
            // Целая часть результата деления
            int result;
            // Остаток деления
            int rest;

            if (isBook)
            {
                #region Пересортировка для книжного формата
                // Для книжного формата может быть не больше трех колонок
                if (columnCount > 3)
                    columnCount = 3;
                // Количество строк
                rowCount = GetHeightCoeff(lastPosition, columnCount);
                // Создание пустого списка блоков подписей
                sortList = CreateEmptySignBlockList(3, rowCount);

                switch (columnCount)
                {
                    case 1:
                        // Запись идет в каждую третью ячейку
                        foreach (var block in list)
                        {
                            position = block.position;
                            index = 3 * position;
                            sortList[index - 1] = block;
                        }
                        break;
                    case 2: 
                        // Нечетные записываются в каждую первую ячейку,
                        // четные - в каждую третью
                        foreach (var block in list)
                        {
                            position = block.position;
                            result = (int)position / 2;
                            rest = position - result * 2;
                            switch (rest)
                            {
                                case 0: index = 3 * result; sortList[index - 1] = block; break;
                                case 1: index = 3 * result + 1; sortList[index - 1] = block; break;
                            }
                        }
                        break;
                    case 3: 
                        foreach(var block in list)
                        {
                            position = block.position;
                            sortList[position - 1] = block;
                        }
                        break;
                }
                #endregion
            }
            else
            {
                #region Пересортировка для альбомного формата
                // Для альбомного формата может быть не больше пяти колонок
                if (columnCount > 5)
                    columnCount = 5;
                // Количество строк
                rowCount = GetHeightCoeff(lastPosition, columnCount);
                // Создание пустого списка блоков подписей
                sortList = CreateEmptySignBlockList(5, rowCount);

                switch (columnCount)
                {
                    case 1: 
                        // Запись идет в каждую пятую ячейку
                        foreach (var block in list)
                        {
                            position = block.position;
                            index = 5 * position;
                            sortList[index - 1] = block;
                        }
                        break;
                    case 2:
                        // Нечетные записываются в каждую первую ячейку,
                        // четные - в каждую пятую
                        foreach (var block in list)
                        {
                            position = block.position;
                            result = (int)(position / 2);
                            rest = position - result * 2;
                            switch(rest)
                            {
                                case 0: index = 5 * result; sortList[index - 1] = block; break;
                                case 1: index = 5 * result + 1; sortList[index - 1] = block; break;
                            }
                        }
                        break;
                    case 3: 
                        foreach (var block in list)
                        {
                            position = block.position;
                            result = (int)(position / 3);
                            rest = position - result * 3;
                            switch (rest)
                            {
                                case 0: index = 5 * result; sortList[index - 1] = block; break;
                                case 1: index = 5 * result + 1; sortList[index - 1] = block; break;
                                case 2: index = 5 * result + 3; sortList[index - 1] = block; break;
                            }                           
                        }
                        break;
                    case 4: 
                        foreach (var block in list)
                        {
                            position = block.position;
                            result = (int)(position / 4);
                            rest = position - result * 4;
                            switch (rest)
                            {
                                case 0: index = 5 * result; sortList[index - 1] = block; break;
                                case 1: index = 5 * result + 1; sortList[index - 1] = block; break;
                                case 2: index = 5 * result + 2; sortList[index - 1] = block; break;
                                case 3: index = 5 * result + 4; sortList[index - 1] = block; break;
                            }
                        }
                        break;
                    case 5: 
                        foreach(var block in list)
                        {
                            position = block.position;
                            sortList[position - 1] = block;
                        }
                        break;
                }
                #endregion
            }

//            sortList = list;

            return sortList;
        }
        // Создание пустого списка SignBlock
        private List<SignBlock> CreateEmptySignBlockList(int columnCount, int rowCount)
        {
            List<SignBlock> list = new List<SignBlock>();

            for (int i = 1; i <= columnCount * rowCount; i++)
                list.Add(null);

            return list;
        }
        // Настройки страницы Excel
        private void PageSetup(Xls xls, bool isBook)
        {
            xls.Worksheet.Cells.RowHeight = 15;
            PageSetup pageSetup = xls.Worksheet.PageSetup;
            pageSetup.PaperSize = XlPaperSize.xlPaperA4;
            if (isBook)
            {
                xls[1, 1, 50, 1].Cells.ColumnWidth = 1.0;
                pageSetup.Orientation = XlPageOrientation.xlPortrait;
            }
            else
            {
                xls[1, 1, 76, 1].Cells.ColumnWidth = 1.0;
                pageSetup.Orientation = XlPageOrientation.xlLandscape;
            }
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 500;
            pageSetup.TopMargin = xls.Application.CentimetersToPoints(1.9);
            pageSetup.BottomMargin = xls.Application.CentimetersToPoints(1.9);
            pageSetup.LeftMargin = xls.Application.CentimetersToPoints(1.8);
            pageSetup.RightMargin = xls.Application.CentimetersToPoints(1.8);
            pageSetup.HeaderMargin = xls.Application.CentimetersToPoints(0.8);
            pageSetup.FooterMargin = xls.Application.CentimetersToPoints(0.8);
        }
    }
}
