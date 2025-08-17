using System;


namespace CertificateSheetReport
{
    public class Guids
    {
        public class NomenclatureParameters
        {
            public static readonly Guid Denotation = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");  //Наименование
            public static readonly Guid Name = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");        //Обозначение
        }

        public static readonly Guid DocToFileLink = new Guid("9eda1479-c00d-4d28-bd8e-0d592c873303");  //Связь документов с файлами
        public static readonly Guid UserShortName = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");  //Короткое имя пользователя

    }
}
