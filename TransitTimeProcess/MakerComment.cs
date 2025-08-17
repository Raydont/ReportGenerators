using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Procedures;
using TFlex.DOCs.Model.References.Users;

namespace TransitTimeProcess
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

        public TimeSpan DurationState = TimeSpan.Zero;

       // public string DurationString = string.Empty;

        public MakerComment(ReferenceObject maker, string comment)
        {
            Maker = maker;
            Comment = comment;
        }

        public MakerComment(KeyValuePair<ReferenceObject, List<ReferenceObject>> stateAction, List<DateTime> datesCreation, byte workType)
        {
            var errorId = 0;
            try
            {
                var actionMaker = stateAction.Value.OrderBy(t => t.SystemFields.CreationDate).Last().GetObjects(Guids.MakersProcessActions);
                errorId = 1;
                if (actionMaker.Count != 0)
                {
                    errorId = 2;
                    Maker = actionMaker.Select(t => (UserReferenceObject)t.GetObject(Guids.LinkMakerToUser)).ToList().OrderByDescending(t => t.SystemFields.CreationDate).FirstOrDefault();
                    Comment = string.Join("\r\n", actionMaker.FirstOrDefault().GetObjects(Guids.CommentsMakersGuid)
                        .Where(t => !string.IsNullOrWhiteSpace(t[Guids.TextCommentGuid].GetString()))
                        .Select(t => t[Guids.TextCommentGuid].GetString().Trim()).ToList());

                    errorId = 2;
                }
                else
                {
                    Maker = stateAction.Key.GetObjects(Guids.MakerStateGuid).FirstOrDefault();

                    errorId = 3;
                }
                WorkType = workType;
                errorId = 4;
                CurrentState = stateAction.Key.ToString();
                errorId = 5;
                DatesCreation.AddRange(datesCreation);
                errorId = 6;
                //  DurationState = TimeSpan.Parse(stateAction.Key[Guids.DurationPlanState].GetString());
                DateTime duration;
                if (stateAction.Key.Class.IsInherit(Guids.TypeStateWorkGuid))
                {
                    duration = ((WorkObject)stateAction.Key).Duration.Value.NowWithOffset;
                }
                else
                {
                    duration = ((AgreementObject)stateAction.Key).Duration.Value.NowWithOffset;
                }
                errorId = 7;
                if (duration != DateTime.MinValue)
                {
                    errorId = 8;
                    DurationState = duration - DateTime.Now;

                    //  DurationString = TermToString(DurationState);

                }
            }
            catch
            {
                MessageBox.Show("Некорректно считались данные у состояния " + stateAction.ToString() + " код ошибки 3." + errorId ,"Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                var actions = stateAction.Value.OrderBy(t => t[Guids.StartTime].GetDateTime());

                foreach (var action in actions)
                {
                    Actions.Add(new ActionParameters(action, stateAction.Key));
                }
            }
            catch
            {
                MessageBox.Show("Некорректно считались данные у состояния " + stateAction.ToString() + " код ошибки 4." + errorId, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string TermToString(TimeSpan term)
        {
            string termToString = string.Empty; 

            if (term.Days < 31 && term.Days >= 1)
                termToString = term.Days + " д.";
            //if (term.Hours < 60 && term.Hours >= 1)
            //    termToString += " " + term.Hours + " ч.";
            //if (term.Minutes < 60 && term.Minutes >= 1)
            //    termToString += " " + term.Minutes + " мин.";
            //if (term.Seconds < 60 && term.Seconds >= 1)
            //    termToString += " " + term.Seconds + " сек.";
            if (termToString != string.Empty)
                termToString += " ";

            if (term.Hours > 0 || term.Minutes > 0 || term.Seconds > 0)
            {
                termToString += term.Hours + ":" + term.Minutes + ":" + term.Seconds;
            }

            return termToString;
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

            if (Maker.Class.IsInherit(Guids.UserTypeGuid))
                result += Maker[Guids.UserShortNameGuid].Value.ToString();
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

        public ActionParameters(ReferenceObject actionObject, ReferenceObject stateObject)
        {
            RefObject = actionObject;
            StartTime = actionObject[Guids.StartTime].GetDateTime();
            EndTime = actionObject[Guids.EndTime].GetDateTime();

            var actionMaker = actionObject.GetObjects(Guids.MakersProcessActions);

            if (actionMaker.Count != 0)
            {
                MakerName = string.Join("\r\n", actionMaker.Select(t => ((UserReferenceObject)t.GetObject(Guids.LinkMakerToUser))[Guids.UserShortNameGuid].GetString()).ToList());
                Comment = string.Join("\r\n", actionMaker.Select(a => string.Join("\r\n", a.GetObjects(Guids.CommentsMakersGuid)
                    .Where(t => !string.IsNullOrWhiteSpace(t[Guids.TextCommentGuid].GetString()))
                    .Select(t => t[Guids.TextCommentGuid].GetString().Trim()).ToList())));
            }
            else
            {
                MakerName = string.Join("\r\n", stateObject.GetObjects(Guids.MakerStateGuid).Select(t=>t.ToString()));
            }          
        }
    }
}
