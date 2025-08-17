using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableCheckInstallationReport
{
    public class ReportParameters
    {
        public int X;
        public int Y;
        public string DenotationDevice;  // Обозначение документа
        public List<string> FileData = new List<string>();          //Первоначальные данные из загружаемого файла
        public List<string> FileDataNet = new List<string>();
        public string FirstUse;          //Первичная применяемость
        public string newFileName;

        public string SignHeader1;
        public string SignHeader2;

        public DateTime SignData1;           //Дата1 подписи     
        public DateTime SignData2;           //Дата2 подписи 

        public string SignName1;
        public string SignName2;

        public string AuthorBy;          // Разраю.
        public string MusteredBy;          // Пров.
        public string NControlBy;          // Н.Контр.
        public string ApprovedBy;          // Утв.

        public DateTime AuthorByDate;         
        public DateTime MusteredByDate;
        public DateTime NControlByDate;
        public DateTime ApprovedByDate;

        public bool AuthorCheck;
        public bool MusteredBCheck;
        public bool NControlByCheck;
        public bool ApprovedByCheck;
        public bool DateStamp1;
        public bool DateStamp2;

        public char Litera1;
        public char Litera2;
        public char Litera3;
        public string NumberSol;
        public string Code;
        public bool PageUD;

        public bool MPP = true;

    }
}

