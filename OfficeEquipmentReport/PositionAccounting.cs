using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;

namespace OfficeEquipmentReport
{
    public class PosAccountingForReport
    {
        public string FIO;
        public string NamePos;
        public string Room;
        public string Type;

        public PosAccountingForReport(string fio, string namePos, string room, string type)
        {
            FIO = fio;
            NamePos = namePos;
            Room = room;
            Type = type;
        }
    }

    public class PositionAccounting
    {
        public string Name = string.Empty;              //Наименование
        public string Remarks;           //Примечание
        public DateTime DeliveryDate;    //Дата поставки
        public DateTime WriteOffDate;    //Дата списания
        public int Status;            //Статус
        public Employee User = new Employee();
        public Room Place = new Room();
        public string Type;
        public ReferenceObject RefPosAc;

        public PositionAccounting(ReferenceObject refObject/*, ReportParameters reportParameters*/)
        {

            int errorId = 0;
            try
            {
                RefPosAc = refObject;
                Name = refObject[Guids.PositionAccountingParameters.NameGuid].Value.ToString();
                errorId = 1;
                Remarks = refObject[Guids.PositionAccountingParameters.RemarksGuid].Value.ToString();
                errorId = 2;
                DeliveryDate = Convert.ToDateTime(refObject[Guids.PositionAccountingParameters.DeliveryDateGuid].Value).Date;
                errorId = 3;
                WriteOffDate = Convert.ToDateTime(refObject[Guids.PositionAccountingParameters.WriteOffDateGuid].Value).Date;
                errorId = 4;
                Status = Convert.ToInt32(refObject[Guids.PositionAccountingParameters.StatusGuid].Value);
                errorId = 5;
                try
                {
                    Place = new Room(refObject.GetObject(Guids.PositionAccountingParameters.AccountPosRoomLinkGuid));
                }
                catch
                {
                    Place = new Room();
                }
                errorId = 7;
                Type = refObject.Class.ToString();
                errorId = 8;

            }
            catch
            {
                MessageBox.Show("Ошибка у объектв " + refObject + ". Код ошибки " + errorId + ". Обратитесь в отдел 911.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public PositionAccounting(PositionAccounting pos, Employee employee)
        {
            Name = pos.Name;
            Remarks = pos.Remarks;
            DeliveryDate = pos.DeliveryDate;
            WriteOffDate = pos.WriteOffDate;
            Status = pos.Status;
            Place = pos.Place;

            if (Name.ToLower().Trim().Contains("мфу"))
                Type = "МФУ";
            else
                Type = pos.Type;

            User = employee;
        }
    }

    public class Room
    {
        public int WorkSpace;    //Количество рабочих мест
        public string Building;  //Корпус 
        public string Name;      //Наименование

        public Room(ReferenceObject refObject)
        {
            WorkSpace = Convert.ToInt32(refObject[Guids.RoomParameters.WorkSpaceCountGuid].Value);
            Building = refObject[Guids.RoomParameters.BuildingGuid].Value.ToString();
            Name = refObject[Guids.RoomParameters.NameGuid].Value.ToString();
        }

        public Room()
        {
        }
     }

    public class Employee
    {
        public string Name;     //ФИО
        public int Status;   //Статус
        public List<PositionAccounting> PosAccounting = new List<PositionAccounting>();
        public ReferenceObject User;

        public Employee(ReferenceObject refObject)//, ReportParameters reportParameters)
        {

            Name = refObject[Guids.EmployeeParameters.NameGuid].Value.ToString();
            Status = Convert.ToInt32(refObject[Guids.EmployeeParameters.StatusGuid].Value);
            User = refObject.GetObject(Guids.EmployeeParameters.LinkToUserGuid);
            foreach (var pos in refObject.GetObjects(Guids.EmployeeParameters.LinkToPositionsAccountingGuid).ToList())
            {
                PosAccounting.Add(new PositionAccounting(pos));
            }
        }
        public Employee()
        {
        }
    }

    public class Department
    {
        public string Name;
        public List<Employee> Employees = new List<Employee>();

        public Department(ReferenceObject refObject)
        {
            Name = refObject[Guids.DepartmentParameters.NameGuid].Value.ToString();
            foreach (var employee in refObject.GetObjects(Guids.DepartmentParameters.LinkToEmployeeGuid).ToList())
            {
                Employees.Add(new Employee(employee));
            }
        }    
    }
}
