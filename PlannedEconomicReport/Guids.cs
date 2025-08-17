using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PlannedEconomicReport
{
    public static class Guids
    {
        public static class UnprovenTime
        {
            public static readonly Guid RefereceGuid = new Guid("1a22decc-996e-464e-b4a4-1ed55eaf164f");

            public static class Types
            {
                public static readonly Guid Department = new Guid("6e2e9a64-c375-4434-8262-a6907b890008");
                public static readonly Guid Leave = new Guid("aaa0da55-4a77-4937-8b88-78917bfdd86d");
                public static readonly Guid SickLeave = new Guid("23777154-ae47-41b5-a7d6-5781d1e222e4");
            }

            public static class Fields
            {
                public static readonly Guid StartDate = new Guid("9fa9cd96-7058-4e39-b83d-a1f2a052724b");
                public static readonly Guid EndDate = new Guid("0567f6f0-8a07-4f8b-b637-ff694eb3cf77");
                public static readonly Guid Name = new Guid("36f0f811-405e-4d2e-8a6a-4fa61c174bd1");
            }

            public static class LeaveFields
            {
                public static readonly Guid DocNumber = new Guid("2290df70-0d68-4c6c-93e3-ad25da55d886");
                public static readonly Guid Type = new Guid("955de103-f4aa-4052-9361-85c85144a254");
                public static readonly Guid ExecutionDate = new Guid("e220b4c3-f736-4a25-aa38-bde2ba0dcbdd");
                public static readonly Guid DayCount = new Guid("5e2c4fb7-7496-47c3-9cd6-419acad57a03");
            }

            public static class Links
            {
                public static readonly Guid Worker = new Guid("64de707d-b767-40ba-a4c0-751620e08560");
            }
        }

        public static class AdditionalPosts
        {
            public static readonly Guid RefereceGuid = new Guid("f42b6350-5efb-4a50-a248-f8c50e5c136a");

            public static class Fields
            {
                public static readonly Guid Name = new Guid("6420aa65-4f14-48a1-9ed6-ae2e9737e416");
            }

            public static class Types
            {
                public static readonly Guid AdditionalPost = new Guid("ef1fed7d-a4e2-4ce9-9d96-54c33c57a1d3");
            }
        }
        public static class Cadres
        {
            public static readonly Guid RefereceGuid = new Guid("22f2bff6-3863-43bb-b9d6-639aeacac81e");

            public static class Fields
            {
                public static readonly Guid Name = new Guid("81a5b4b8-8621-4330-a7d9-42bb2558c4c9");
            }

            public static class WorkerFields
            {
                public static readonly Guid Post = new Guid("faa33fbf-96c0-435b-83f1-a2c5770f3e41");
                public static readonly Guid Pay = new Guid("ac5e2409-30e6-48b7-98d1-14a160c3693b");
                public static readonly Guid TimeBoardNumber = new Guid("35641b21-3425-4142-b5c6-08f84617b58f");
            }

            public static class Types
            {
                public static readonly Guid Department = new Guid("24e75318-6cd5-45fb-9947-a325ffd34792");
                public static readonly Guid Worker = new Guid("a6ce9468-d6f7-4a39-8564-050b98706b09");
            }

            public static class Links
            {
                public static readonly Guid User = new Guid("650d1f1a-ce38-4fb6-8668-88c84c7855e5");
                public static readonly Guid DayOffRequests = new Guid("975369c2-883d-4051-9748-a5740cd0fc40");
                public static readonly Guid UnprovenTimes = new Guid("64de707d-b767-40ba-a4c0-751620e08560");
                public static readonly Guid Leaves = new Guid("55112d87-2873-4f15-adf0-fe7a11dafa59");
                public static readonly Guid AdditionalPosts = new Guid("85d4cc54-6f3b-4b66-bc8c-667da23cf886");
            }
        }



    }
}
