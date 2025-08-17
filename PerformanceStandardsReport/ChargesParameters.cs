using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PerformanceStandardsReport
{
    public class ChargesParameters
    {
        public static class ReferenceGuid
        {
            public static readonly Guid Workers = new Guid("7776ab19-8f42-470c-8863-c1a17f2394f4");
            public static readonly Guid Charges = new Guid("858e20f8-b25d-45fb-9844-2c170a987465");
            public static readonly Guid CostHoursOverFullfilledTime = new Guid("5b254aa4-c1a3-4283-8537-8e33134103d6");
            public static readonly Guid Users = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        }

        public static class ReferenceTypesGuid
        {
            public static readonly Guid Worker = new Guid("8f24dc2f-4a8c-4215-b766-6d22d5609acd");
            public static readonly Guid Department = new Guid("c46d5686-5776-4d37-8ccc-a3f747468c5c");

            public static readonly Guid Charge = new Guid("3951ea5e-0677-46b9-b201-c696e0a4b629");
            public static readonly Guid Folder = new Guid("5b486977-80c9-400b-8792-7b68512612ba");

            public static readonly Guid CostHoursOverFullfilledTime = new Guid("1fe9ddde-6719-46c1-8562-660f9e04f3b3");
        }
        public static class WorkerGuid
        {
            public static readonly Guid Name = new Guid("e307df8e-47d1-4214-9d5b-3dbdcbf118ea");
            public static readonly Guid Payment = new Guid("198df15a-1223-462c-8bfa-2af763983313");
            public static readonly Guid Profession = new Guid("4dea7de5-9e7c-40e0-9793-40459d1f4238");
            public static readonly Guid Category = new Guid("525c8cab-df1e-4c4f-9941-0183a847ff1e");
            public static readonly Guid TimeBoardNumber = new Guid("43440925-7eff-44eb-bf39-fbc2217a7ed9");
            public static readonly Guid WageRate = new Guid("f0504b63-0dc0-49de-a6d7-5d56665956e1");
            public static readonly Guid PersentBasicAwards = new Guid("f0eca4dd-c647-415f-84cf-faceebef0fbd");
            public static readonly Guid PersentAdditionalAwards = new Guid("380520dd-291f-407d-bc21-16e5581e06d7");

            public static readonly Guid LinkToOverFullfilledTime = new Guid("7c51174f-de88-44eb-b181-e388d556f751");
        }

        public static class CostHoursOverFullfilledTimeGuid
        {
            public static readonly Guid CountHoursListObjects = new Guid("5e3cdf8d-33b6-4a2a-984b-a529a5b1f217");
            public static readonly Guid Name = new Guid("e2d37366-36eb-4c42-b9d1-23ceabb1240d");
            public static readonly Guid CountHours = new Guid("23ecfeb9-b464-438b-81e5-081b7a4a3377");
            public static readonly Guid CostHours = new Guid("9eefd1d7-cde1-43e6-afaf-a08f0637c8be");
        }

        public static class ChargesGuid
        {
            public static readonly Guid Name = new Guid("d179afd5-ffb2-4790-aa01-79f0d5c87ecc");

            public static readonly Guid Surcharges = new Guid("93867563-77fe-4619-9733-576745cf1e7a");
            public static readonly Guid AdditionalAward = new Guid("0bb467d0-9d73-4920-87cb-b0699197aa33");
            public static readonly Guid BasicAward = new Guid("d8ba3e6d-b894-420a-aa3c-49d203eae6c4");
            public static readonly Guid FullFilledHours = new Guid("de5445c2-a365-4036-925c-6c12947c8cf6");
            public static readonly Guid ActualFullFilledHours = new Guid("25c0906d-9ddc-4ac7-8aa7-535dfe13fab1");
            public static readonly Guid LinkToWorker = new Guid("21a79c0d-2ccf-431a-a89b-b80fdba232d8");
            public static readonly Guid SurchargeFullFilledHours = new Guid("30c9cacf-0fa6-4a0c-8f5f-c9c93e6e4e4c"); //Отработанное время для доплат
            public static readonly Guid OverTimeFullFilledHours = new Guid("fec712ab-9b40-408d-9757-f5427546f1f1");  //Отработанное время сверхурочно
            public static readonly Guid DaysOffFulFilledHours = new Guid("a31f4767-b8c9-466c-887b-24e61e82512a");    //Отработанное время в выходные - РВ1
            public static readonly Guid DaysOffFulFilledHours2 = new Guid("d8790baf-5a95-43df-9a51-ea90468ee8de");    //Отработанное время в выходные - РВ2
            public static readonly Guid EmploymentPercent = new Guid("17aab739-c2ec-439e-8cf0-dd58a315de88");        //Процент занятости для доплат
            public static readonly Guid DevelopedStandardHours = new Guid("3c74de09-5858-4a5f-ba2b-8f17fdfcedea");   //Выработанные нормочасы
            public static readonly Guid Pupil = new Guid("0553feea-0c97-4d8e-b15e-0ae175c0bc8c");   // Ученик
            public static readonly Guid PartTimeWorker = new Guid("58111e72-437d-42e7-910a-2df5376660dc");  //Совместитель
        }

        public static readonly Guid WorkToCharges = new Guid("aa6c8a9c-85d7-4f8b-b090-d637227b695a");
        public static readonly Guid ShortNameUser = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");
    }
}
