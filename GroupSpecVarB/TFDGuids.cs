using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;

namespace GroupSpecVarB
{
    public class TFDGuids
    {
        public static class NomenclatureObjectTypesId
        {
            public static readonly int Assembly = new NomenclatureReference().Classes.AllClasses.Find("Сборочная единица").Id;
            public static readonly int Detal = new NomenclatureReference().Classes.AllClasses.Find("Деталь").Id;
            public static readonly int OtherDocument = new NomenclatureReference().Classes.AllClasses.Find("Документ").Id;
            public static readonly int Document = new NomenclatureReference().Classes.AllClasses.Find(new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114")).Id;
            public static readonly int Komplekt = new NomenclatureReference().Classes.AllClasses.Find("Комплект").Id;
            public static readonly int Complex = new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Id;
            public static readonly int OtherItem = new NomenclatureReference().Classes.AllClasses.Find("Прочее изделие").Id;
            public static readonly int StandartItem = new NomenclatureReference().Classes.AllClasses.Find("Стандартное изделие").Id;
            public static readonly int Material = new NomenclatureReference().Classes.AllClasses.Find("Материал").Id;
            public static readonly int ComponentProgram = new NomenclatureReference().Classes.AllClasses.Find("Компонент (программы)").Id;
            public static readonly int ComplexProgram = new NomenclatureReference().Classes.AllClasses.Find("Комплекс (программы)").Id;
        }

        public static readonly Guid RefNomenclatureGuid = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");

        public static class Types
        {
            public static readonly Guid Document = new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114");
            public static readonly Guid Detal = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
            public static readonly Guid SborochnayaEdinica = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
            public static readonly Guid Complekt = new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340");
            public static readonly Guid Complex = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");
            public static readonly Guid Izdelie = new Guid("7fa98498-c39c-44fc-bcaa-699b387f7f46");
            public static readonly Guid StandartnoeIzdelie = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");
            public static readonly Guid Material = new Guid("f7f45e16-ceba-4d26-a9af-f099a2e2fca6");
            public static readonly Guid ProcheeIzdelie = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
            public static readonly Guid MaterialGlobus = new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987");
        }

        public static class HierarchyLink
        {

            public static readonly Guid Poziciya = new Guid("ab34ef56-6c68-4e23-a532-dead399b2f2e");
            public static readonly Guid Kolichestvo = new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818");
            public static readonly Guid StandardItemStore = new Guid("d174b609-029d-405b-834c-868076bc2147");
            public static readonly Guid EdinicaIzmereniya = new Guid("530922fa-8490-49c8-b93a-5e604edc1d7d");
            public static readonly Guid Comment = new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce");
            public static readonly Guid Zona = new Guid("1367dda9-7850-4c15-b123-636ab692034c");
        }

        // Параметры объектов Номенклатуры
        public static class ParInfo
        {
            // Наименование
            public static readonly ParameterInfo Name =
                new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(
                    NomenclatureReferenceObject.FieldKeys.Name);

            // Обозначение
            public static readonly ParameterInfo Denotation =
                new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(
                    NomenclatureReferenceObject.FieldKeys.Denotation);

            // Класс объекта
            public static readonly ParameterInfo Class =
                new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(SystemParameterType.Class);

            // ID объекта
            public static readonly ParameterInfo ObjectId =
                new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(SystemParameterType.ObjectId);

            // Формат объекта 
            public static readonly ParameterInfo Format =
                new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(
                    new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c"));

        }

        public static class BomSection
        {
            public static readonly string Documentation = "Документация";
            public static readonly string Assembly = "Сборочные единицы";
            public static readonly string Details = "Детали";
            public static readonly string StdProducts = "Стандартные изделия";
            public static readonly string OtherProducts = "Прочие изделия";
            public static readonly string Materials = "Материалы";
            public static readonly string Komplekts = "Комплекты";
            public static readonly string Complex = "Комплексы";
            public static readonly string Components = "Компоненты";
            public static readonly string ProgramItems = "Программные изделия";

        }
    }
}
