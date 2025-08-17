using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;

namespace CalculationOtherProducts
{
    public class ReportParameters
    {
        public ReferenceObject RootObject;           //Корневой объект
        public bool OtherItems;    //Прочие изделия
        public bool Materials;     //Материалы
        public bool StandartItems; //Стандартные изделия
        public bool Assembly;      //Сборочные единицы
        public bool Details;       //Детали 
        public bool UseCurrentObject; //Разузловывать выделенный объект

        public List<NomSpecificationItems> Devices = new List<NomSpecificationItems>();      //Список загруженных устройств
        public List<NomSpecificationItems> PKIObjects = new List<NomSpecificationItems>();   //Список загруженных объектов типа Прочих изделий
    }
}
