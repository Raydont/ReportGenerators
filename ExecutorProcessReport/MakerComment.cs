using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Procedures;
using TFlex.DOCs.Model.References.Users;

namespace ExecutorProcessReport
{
    public class MakerComment
    {
        public ReferenceObject Maker; //Исполнитель
        public string SignatureType = string.Empty;
        public string Comment = string.Empty; //Комментарий
        public byte WorkType; //Тип работы 1 - В работе, 2 - На доработке, 3 - Подпись установлена
        public string CurrentState = string.Empty; //Текущее состояние
        public List<DateTime> DatesCreation = new List<DateTime>();
        public List<ActionParameters> Actions = new List<ActionParameters>();
        public ReferenceObject RKK;

        public TimeSpan DurationState = TimeSpan.Zero;

        // public string DurationString = string.Empty;

        public MakerComment(ReferenceObject maker, string comment)
        {
            Maker = maker;
            Comment = comment;
        }

        public MakerComment(KeyValuePair<ReferenceObject, List<ReferenceObject>> stateAction, List<DateTime> datesCreation, byte workType)
        {
            var actionMaker = stateAction.Value.OrderBy(t => t.SystemFields.CreationDate).Last().GetObjects(Guids.ДействияПроцессов.СписокОбъектов.Исполнители.Id);

            if (actionMaker.Count != 0)
            {
                Maker = actionMaker.Select(t => (UserReferenceObject)t.GetObject(Guids.ДействияПроцессов.СписокОбъектов.Исполнители.Связи.Пользователь)).ToList().OrderByDescending(t => t.SystemFields.CreationDate).FirstOrDefault();
                Comment = string.Join("\r\n", actionMaker.FirstOrDefault().GetObjects(Guids.ДействияПроцессов.СписокОбъектов.Исполнители.СпискиОбъектов.Комментарии.Id)
                    .Where(t => !string.IsNullOrWhiteSpace(t[Guids.ДействияПроцессов.СписокОбъектов.Исполнители.СпискиОбъектов.Комментарии.Поля.Текст].GetString()))
                    .Select(t => t[Guids.ДействияПроцессов.СписокОбъектов.Исполнители.СпискиОбъектов.Комментарии.Поля.Текст].GetString().Trim()).ToList());
            }
            else
            {
                Maker = stateAction.Key.GetObjects(Guids.Процедуры.СписокОбъектов.Состояния.Связи.Исполнители).FirstOrDefault();
            }
            WorkType = workType;
            CurrentState = stateAction.Key.ToString();
            DatesCreation.AddRange(datesCreation);

            //  DurationState = TimeSpan.Parse(stateAction.Key[Guids.DurationPlanState].GetString());
            try
            {
                var duration = ((WorkObject)stateAction.Key).Duration.Value.NowWithOffset;
                if (duration != DateTime.MinValue)
                {
                    DurationState = duration - DateTime.Now;
                }
            }
            catch
            {

            }
            var actions = stateAction.Value.OrderBy(t => t[Guids.Процедуры.СписокОбъектов.Состояния.Поля.ВремяНачала].GetDateTime());

            foreach (var action in actions)
            {
                Actions.Add(new ActionParameters(action, stateAction.Key));
            }

        }
  

        public MakerComment(string signatureType, ReferenceObject maker, string comment, DateTime dateCreation, byte workType)
        {
            SignatureType = signatureType;
            Maker = maker;
            Comment = comment;
            DatesCreation.Add(dateCreation);
            WorkType = workType;
        }

        public override string ToString()
        {
            string result = string.Empty;
            if (CurrentState.Trim() != string.Empty)
                result = CurrentState + ". ";

            if (SignatureType.Trim() != string.Empty)
                result = SignatureType + " - ";

            if (Maker.Class.IsInherit(Guids.ГруппыИПользователи.Типы.Пользователь))
                result += Maker[Guids.ГруппыИПользователи.Поля.КороткоеИмя].Value.ToString();
            else
                result += Maker.ToString();

            if (Comment.Trim() != string.Empty)
            {
                result += ":\r\n" + Comment.Trim();
            }

            result += ":\r\n" + DurationState.ToString().Trim();


            return result;
        }
    }

    public class ActionParameters
    {
        public ReferenceObject RefObject;
        public string Comment = string.Empty;
        public DateTime StartTime;
        public DateTime EndTime;
        public string MakerName;
        public ReferenceObject RKK;

        public ActionParameters(ReferenceObject actionObject, ReferenceObject stateObject)
        {
            RefObject = actionObject;
            StartTime = actionObject[Guids.Процедуры.СписокОбъектов.Состояния.Поля.ВремяНачала].GetDateTime();
            EndTime = actionObject[Guids.Процедуры.СписокОбъектов.Состояния.Поля.ВремяЗавершения].GetDateTime();

            var actionMaker = actionObject.GetObjects(Guids.ДействияПроцессов.СписокОбъектов.Исполнители.Id);

            if (actionMaker.Count != 0)
            {
                MakerName = string.Join("\r\n", actionMaker
                    .Where(t => ((UserReferenceObject)t.GetObject(Guids.ДействияПроцессов.СписокОбъектов.Исполнители.Связи.Пользователь))[Guids.ГруппыИПользователи.Поля.КороткоеИмя].GetString().Trim() != string.Empty)
                    .Select(t => ((UserReferenceObject)t.GetObject(Guids.ДействияПроцессов.СписокОбъектов.Исполнители.Связи.Пользователь))[Guids.ГруппыИПользователи.Поля.КороткоеИмя].GetString())
                    .Distinct().ToList());
                Comment = string.Join("\r\n", actionMaker.Select(a => string.Join("\r\n", a.GetObjects(Guids.ДействияПроцессов.СписокОбъектов.Исполнители.СпискиОбъектов.Комментарии.Id)
                    .Where(t => !string.IsNullOrWhiteSpace(t[Guids.ДействияПроцессов.СписокОбъектов.Исполнители.СпискиОбъектов.Комментарии.Поля.Текст].GetString()))
                    .Select(t => t[Guids.ДействияПроцессов.СписокОбъектов.Исполнители.СпискиОбъектов.Комментарии.Поля.Текст].GetString().Trim()).ToList())));
            }
            else
            {
                MakerName = string.Join("\r\n", stateObject.GetObjects(Guids.Процедуры.СписокОбъектов.Состояния.Связи.Исполнители).Select(t => t.ToString()));
            }
        }
    }
}
