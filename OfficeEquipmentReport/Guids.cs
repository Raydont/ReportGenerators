using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeEquipmentReport
{
    public class Guids
    {
        public static class ReferencesGuid
        {
            public static readonly Guid PostitionAccountingGuid = new Guid("07b3ee6b-2526-4cd2-9787-ad417b72678c"); //Справочник позиции учета
            public static readonly Guid DepartmentsGuid = new Guid("7f41deed-e6dd-459b-a357-cec65d2d3193");         //Подразделения
            public static readonly Guid RoomsGuid = new Guid("c4544eae-e1d6-4787-ad35-71d458649a46");               //Комнаты
            public static readonly Guid EmployeeGuid = new Guid("c1a60d47-a82e-48f0-9de4-b13472826671");            //Сотрудники
            public static readonly Guid GroupAndUserGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");        //Группы и пользователи
        }

        public static class GroupAndUser
        {
            public static readonly Guid BlockingUserTypeGuid = new Guid("43bcafc5-9f4f-48ef-acb2-2c0177ad1a7c");    //Отключенный пользователь
        }

        public class PositionAccountingTypes
        {
            public static readonly Guid ComputerTypeGuid = new Guid("e5607a6d-c97c-4835-812b-9c02b246e8ae");         //Системный блок
            public static readonly Guid MonitorTypeGuid = new Guid("1bae6108-fc9b-478f-abe3-6698f1567674");          //Монитор
            public static readonly Guid PrinterTypeGuid = new Guid("608d0af5-1c6d-4fec-a1c4-4325387bdf3c");          //Принтер
            public static readonly Guid ScannerTypeGuid = new Guid("bbfa3bac-b4d8-436e-86db-0d6151a496cd");          //Сканер
            public static readonly Guid StorageTypeGuid = new Guid("a66094c5-a101-4def-aaf3-ae6daa6c638a");          //Накопитель
            public static readonly Guid OtherDevicesTypeGuid = new Guid("3ebd547d-a817-4869-a1a5-65184d99a23e");     //Прочие устройства
            public static readonly Guid NotebookTypeGuid = new Guid("3dc7ce8d-d6e2-415c-97db-e6831c5c7f46");         //Ноутбук
        }
        public class PositionAccountingParameters
        {
            public static readonly Guid NameGuid = new Guid("2f7b4613-fae2-484b-9989-4fc528bfab10");                 //Наименование
            public static readonly Guid RemarksGuid = new Guid("037d8938-66d5-4808-affb-d66b85d91bc5");              //Примечание
            public static readonly Guid DeliveryDateGuid = new Guid("d68303a0-0666-4aae-a6de-6a4c186a2030");         //Дата поставки
            public static readonly Guid WriteOffDateGuid = new Guid("d68303a0-0666-4aae-a6de-6a4c186a2030");         //дата списания
            public static readonly Guid StatusGuid = new Guid("fc454415-27a2-4cea-92d0-b27a794d0c36");               //Статус
            public static readonly Guid AccountPosRoomLinkGuid = new Guid("f25acebf-bdeb-4520-b991-e88d103ec603");   //Связь системного блока с комнатой
            public static readonly Guid EmployeeAccountLinkGuid = new Guid("f1b3f0b1-b18b-476f-9924-11d258ffbd96");  //Связь системного блока с сотрудником
        }
        public static class RoomsTypes
        {
            public static readonly Guid RoomTypeGuid = new Guid("da2de9fb-f55e-4f0b-a5ee-4dd66afa8386");    //тип сотрудник 
        }

        public static class EmpoyeeTypes
        {
            public static readonly Guid EmployeeTypeGuid = new Guid("47ec2b8b-5f42-4c0d-ab6a-ff3f831f4f39"); //Тип сотрудник
        }

        public static class DepartmentTypes
        {
            public static readonly Guid DepartmentTypeGuid = new Guid("c614a258-911d-4a21-848e-593bd641239f"); //Тип подразделение
        }

        public static class ComputerParameters
        {
            public static readonly Guid HDDGuid = new Guid("5df5ffdf-66b1-4e30-95d5-2703962bee6b");                  //Жесткий диск
            public static readonly Guid VideoCardGuid = new Guid("e2f23600-fbeb-4f1c-97cd-76b94e16007b");            //Видео карта
            public static readonly Guid MotherboardGuid = new Guid("f3ef7644-568d-4930-a8c8-ed3bc93f0e98");          //Материнская плата
            public static readonly Guid RAMGuid = new Guid("9034377e-f27f-4d37-9221-107fa547815e");                  //Оперативная память
            public static readonly Guid OpticalDriveGuid = new Guid("a2cdc82a-ebbc-4681-a7fa-f2ac117a076d");         //Оптический привод
            public static readonly Guid ProcessorGuid = new Guid("9d79f953-5e77-4ed6-a8a3-cbbadc3ae58a");            //Процессор          
        }

        public static class RoomParameters
        {
            public static readonly Guid NameGuid = new Guid("cefb4422-56ae-4484-9cc8-4e228dd93a13");            //Наименвоание
            public static readonly Guid BuildingGuid = new Guid("b2e9383b-bf4c-4ce1-9a50-00eff14466f1");        //Корпус
            public static readonly Guid WorkSpaceCountGuid = new Guid("aebdac93-4693-43ba-a3f3-c98654402ac7");  //Количество рабочих мест
            public static readonly Guid LinkToEmployeeGuid = new Guid("202129b6-5559-41e1-ba0a-6ca479360239");  //связь с сотрудником
        }

        public static class PrinterParameters
        {
            public static readonly Guid TypeGuid = new Guid("a27af888-83b8-4eb8-b63a-5d7bfab4f4a2");    //Тип
            public static readonly Guid FormatGuid = new Guid("2765fd56-471d-4cf7-a10e-5584881b74a7");  //Формат
        }

        public static class ScannerParameters
        {
            public static readonly Guid FormatGuid = new Guid("bbfa3bac-b4d8-436e-86db-0d6151a496cd"); //Формат
        }

        public static class StorageParameters
        {
            public static readonly Guid VolumeGuid = new Guid("3c143497-a305-49ca-97d3-9ffbb298c439");  //Объем
            public static readonly Guid TypeGuid = new Guid("87dff65e-8f6f-418a-8201-a32a86db2088");    //Тип
        }

        public static class EmployeeParameters
        {
            public static readonly Guid NameGuid = new Guid("5b43a331-446a-45e0-9e54-3e36332462e2");                      //ФИО
            public static readonly Guid StatusGuid = new Guid("623695ff-f2f5-4aa5-93b8-300e7754e944");                    //Статус
            public static readonly Guid LinkToDepartmentGuid = new Guid("974adbca-af61-49dd-bf3b-fe510f4a2693");          //Связь с подразделением
            public static readonly Guid LinkToUserGuid = new Guid("fad9978d-a939-4491-9404-e69831457057");                //Связь с пользователем
            public static readonly Guid LinkToPositionsAccountingGuid = new Guid("f1b3f0b1-b18b-476f-9924-11d258ffbd96"); //Связь с позицией учета
            public static readonly Guid LinkToRoomsGuid = new Guid("202129b6-5559-41e1-ba0a-6ca479360239");               //Связь с комнатой
        }

        public static class DepartmentParameters
        {
            public static readonly Guid NameGuid = new Guid("8a9c941d-4c3a-4d7f-ae54-4348436bce03");                //Наименование
            public static readonly Guid CountRoomsGuid = new Guid("6a3ea7e7-6153-4446-93ea-d7eebb82efbd");          //Количество комнат
            public static readonly Guid CountEmployeeGuid = new Guid("12c45b73-5602-4601-aa59-8d8058003cb3");       //Количество сотрудников
            public static readonly Guid LinkToRoomsGuid = new Guid("9fc786ac-e309-4012-aa54-4afc382f9590");         //Связь с комнатами
            public static readonly Guid LinkToEmployeeGuid = new Guid("541343e0-abc5-40a0-aa48-7cb1e123b4ca");      //Связь с сотрудниками
        }

        public static class OtherDevicesParameters
        {
            public static readonly Guid TypeGuid = new Guid("204d615c-fcca-4d5b-855c-66e41cebc506"); //Тип
            public static readonly Guid CountGuid = new Guid("10f601e1-98d7-4924-a0ea-681aa6140485"); //Количество
        }
    }
}
