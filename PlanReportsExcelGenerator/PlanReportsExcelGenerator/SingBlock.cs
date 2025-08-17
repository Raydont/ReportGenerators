using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportHelpers
{
    public enum SignType
    {
        /// <summary>
        /// УТВЕРЖДЕНО
        /// </summary>
        Confirmed,
        /// <summary>
        /// СОГЛАСОВАНО 
        /// </summary>
        Agreed
    }

    public enum DepartmentType
    {
        /// <summary>
        /// Отдел
        /// </summary>
        Depart,
        /// <summary>
        /// Бюро
        /// </summary>
        Bureau,
        /// <summary>
        /// Цех
        /// </summary>
        Shop,
        /// <summary>
        /// Подразделение
        /// </summary>
        Division,
        /// <summary>
        /// ВП МО
        /// </summary>
        VPMO,
        /// <summary>
        /// Генеральный директор
        /// </summary>
        GeneralDirector,
        /// <summary>
        /// Главный конструктор
        /// </summary>
        ChiefDesigner,
        /// <summary>
        /// Главный инженер
        /// </summary>
        ChiefEngineer,
        /// <summary>
        /// Зам.генерального директора
        /// </summary>
        DeputyGeneralDirector
    }

    public class SignBlock
    {
        public SignType signType;

        public string GetSignTypeString()
        {
            if (signType == SignType.Confirmed)
                return "УТВЕРЖДАЮ";
            if (signType == SignType.Agreed)
                return "СОГЛАСОВАНО";

            return "";
        }

        public DepartmentType departmentType;

        public string GetDepartmentTypeChiefString()
        {
            if (departmentType == DepartmentType.GeneralDirector)
                return "Генеральный директор";
            if (departmentType == DepartmentType.ChiefDesigner)
                return "Главный конструктор";
            if (departmentType == DepartmentType.ChiefEngineer)
                return "Главный инженер";
            if (departmentType == DepartmentType.DeputyGeneralDirector)
                return "Зам.генерального директора";
            if (departmentType == DepartmentType.VPMO)
                return "Начальник ВП МО";
            if (departmentType == DepartmentType.Depart)
                return "Начальник отдела " + departmentNumber;
            if (departmentType == DepartmentType.Bureau)
                return "Начальник бюро " + departmentNumber;
            if (departmentType == DepartmentType.Shop)
                return "Начальник цеха " + departmentNumber;
            if (departmentType == DepartmentType.Division)
                return "Начальник подразделения " + departmentNumber;

            return "";
        }

        public string departmentNumber = "";
        public string fio;
        public int position;
    }
}
