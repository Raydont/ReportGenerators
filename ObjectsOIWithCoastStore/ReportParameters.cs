using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;

namespace ObjectsOIWithCoastStore
{
    public class ReportParameters
    {
        public bool SelectedDevice = false;
        public bool SelectedStoreObjects = false;
        public bool SelectedDevices = false;
        public bool SelectedNomObjects = false;
        public bool SelectedDevicesWithKoef = false;
        public bool SelectedDevicesDifferentFiles = false;
        public bool ObjectsReferenceFailure = false;
        public string SelectedPath = string.Empty;

        public NomenclatureObject Device = null;
        public List<StoreObject> StoreObjects = new List<StoreObject>();
        public Dictionary<NomenclatureObject, int> Devices = new Dictionary<NomenclatureObject, int>();
        public List<NomObject> NomObjects = new List<NomObject>();
        public List<DeviceWithKoef> DevicesWithKoef = new List<DeviceWithKoef>();
        public List<DeviceWithCountByRow> DevicesByRow = new List<DeviceWithCountByRow>();

        public DateTime BeginPeriod = DateTime.MinValue;
        public DateTime EndPeriod = DateTime.MinValue;
    }

    public class DeviceWithKoef
    {
        public NomenclatureObject NomObj;
        public int Count;
        public double Koef;

        public DeviceWithKoef(NomenclatureObject nomObj, int count, double koef)
        {
            NomObj = nomObj;
            Count = count;
            Koef = koef;
        }
    }

    public class DeviceWithCountByRow
    {
        public NomenclatureObject NomObj;
        public int Count;
        public int Row;

        public DeviceWithCountByRow(NomenclatureObject nomObj, int count, int row)
        {
            NomObj = nomObj;
            Count = count;
            Row = row;
        }
    }
}
