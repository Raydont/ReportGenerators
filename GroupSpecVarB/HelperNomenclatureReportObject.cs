using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References.Nomenclature;

namespace GroupSpecVarB
{
    public enum ReportObjectType
    {
        Uknown,
        Detail,
        Izdelie,
        Complex,
        SborochnayaEdinica,
        Complekt,
        StandartnoeIzdelie,
        Material,

        Document,
        ProcheeIzdelie,
        ProgrammnoeIzdelie,
        Component
    }

    public class HelperNomenclatureReportObject
    {

        public List<int>listNumebrMake = new List<int>(); 

        private  NomenclatureObject _nomenclatureObject;

        public  NomenclatureObject NomenclatureObj
        {
            get 
            { return _nomenclatureObject; }
            set
            {
                _nomenclatureObject = value;

                if (_nomenclatureHierarchyLink != null) return;

                if (_nomenclatureObject.Class.Guid == TFDGuids.Types.Detal)
                {
                    ObjectType = ReportObjectType.Detail;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.Izdelie)
                {
                    ObjectType = ReportObjectType.Izdelie;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.Complex)
                {
                    ObjectType = ReportObjectType.Complex;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.Complekt)
                {
                    ObjectType = ReportObjectType.Complekt;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.SborochnayaEdinica)
                {
                    ObjectType = ReportObjectType.SborochnayaEdinica;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.Material)
                {
                    ObjectType = ReportObjectType.Material;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.MaterialGlobus)
                {
                    ObjectType = ReportObjectType.Material;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.StandartnoeIzdelie)
                {
                    ObjectType = ReportObjectType.StandartnoeIzdelie;
                }
                else if (_nomenclatureObject.Class.Guid == TFDGuids.Types.Document)
                {
                    ObjectType = ReportObjectType.Document;
                }
            }
        }

        private ReportObjectType _objectType = ReportObjectType.Uknown;

        public ReportObjectType ObjectType
        {
            get { return _objectType; }
            private set { _objectType = value; }
        }

        private NomenclatureHierarchyLink _nomenclatureHierarchyLink;

        public NomenclatureHierarchyLink NomenclatureHLink
        {
            get { return _nomenclatureHierarchyLink; }
            set
            {
                _nomenclatureHierarchyLink = value;
                if (_nomenclatureHierarchyLink != null)
                {
                    bomSection =
                        _nomenclatureHierarchyLink.ParameterValues[NomenclatureHierarchyLink.FieldKeys.BomSection].
                            GetString();
                    position =
                        _nomenclatureHierarchyLink.ParameterValues[NomenclatureHierarchyLink.FieldKeys.Position].
                            GetInt32();
                    kolichestvo =
                        _nomenclatureHierarchyLink.ParameterValues[TFDGuids.HierarchyLink.Kolichestvo].GetSingle();
                    comment = _nomenclatureHierarchyLink.ParameterValues[TFDGuids.HierarchyLink.Comment].GetString();
                    zona =
                        (_nomenclatureHierarchyLink.ParameterValues[TFDGuids.HierarchyLink.Zona].GetString() ?? "").Trim();

                  

                    switch (bomSection)
                    {
                        case "Документация":
                            ObjectType = ReportObjectType.Document;
                            break;
                        case "Детали":
                            ObjectType = ReportObjectType.Detail;
                            break;
                        case "Прочие изделия":
                            ObjectType = ReportObjectType.ProcheeIzdelie;
                            break;
                        case "Сборочные единицы":
                            ObjectType = ReportObjectType.SborochnayaEdinica;
                            break;
                        case "Стандартные изделия":
                            ObjectType = ReportObjectType.StandartnoeIzdelie;
                            break;
                        case "Комплексы":
                            ObjectType = ReportObjectType.Complex;
                            break;
                        case "Комплекты":
                            ObjectType = ReportObjectType.Complekt;
                            break;
                        case "Материалы":
                            ObjectType = ReportObjectType.Material;
                            break;
                        case "Программные изделия":
                            ObjectType = ReportObjectType.ProgrammnoeIzdelie;
                            break;
                        case "Компоненты":
                            ObjectType = ReportObjectType.Component;
                            break;
                        default:
                            ObjectType = ReportObjectType.Uknown;
                            break;
                    }
                }

            }
        }

        public HelperNomenclatureReportObject Parent { get; set; }

        public bool Compare (HelperNomenclatureReportObject helperNomObj)
        {
            if (name == helperNomObj.name && denotation == helperNomObj.denotation && objectID == helperNomObj.objectID 
                && format == helperNomObj.format && position == helperNomObj.position && kolichestvo == helperNomObj.kolichestvo
                && bomSection == helperNomObj.bomSection && comment == helperNomObj.comment && zona == helperNomObj.zona)
            {
                return true;
            }
            else return false;
        }
        
        public int position;
        public float kolichestvo;
        public string bomSection;
        public string comment;
        public string zona;
        public bool firstLevel;
        public int numberMake;
        public string _denotation;
        public int nominalOtherItems;
        
        public string _name;
     
        public string numberMakeStr

        {
            get
            {
                if (numberMake < 10)
                    return "0" + numberMake.ToString();
                else 
                    return  numberMake.ToString();
            } 
        }


        public string name
        {
            get { return NomenclatureObj[TFDGuids.ParInfo.Name].Value.ToString(); }
        }
        public string denotation
        {
            get { return NomenclatureObj[TFDGuids.ParInfo.Denotation].Value.ToString(); }
        }
        public int objectID
        {
            get { return Convert.ToInt32(NomenclatureObj[TFDGuids.ParInfo.ObjectId].Value); }
        }

        public int parentID;
        //{
         //   get { return Convert.ToInt32(NomenclatureObj.Parent[TFDGuids.ParInfo.ObjectId]); }
       // }
        public int type
        {
            get { return Convert.ToInt32(NomenclatureObj[TFDGuids.ParInfo.Class].Value); }
        }
        public string format
        {
            get { return NomenclatureObj[TFDGuids.ParInfo.Format].Value.ToString(); }
        }

    }
}
