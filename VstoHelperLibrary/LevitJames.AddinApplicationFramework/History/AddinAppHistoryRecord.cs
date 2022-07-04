// © Copyright 2018 Levit & James, Inc.

using System;
using System.Runtime.Serialization;
using System.Text;
using LevitJames.TextServices;

namespace LevitJames.AddinApplicationFramework
{
    [Serializable]
    public class AddinAppHistoryRecord : AppSerializableBase
    {
        public AddinAppHistoryRecord(int level, string detailText, string itemText)
        {
            DetailText = detailText;
            Level = level;
            ItemText = itemText;
        }

        protected AddinAppHistoryRecord(SerializationInfo info, StreamingContext context) : base(info, context) { }


        // PUBLIC PROPERTIES

        public int Id { get; internal set; }
        public DateTime TimeStamp { get; internal set; }
        public int Level { get; set; }
        public string DetailText { get; set; }
        public string Username { get; internal set; }
        public string ItemText { get; set; } // Comma-delimited list of WtPageItemBase.Id values
        public string AppVersion { get; internal set; }


        public string DisplayText
        {
            get
            {
                const int maxDisplayChars = 50;
                var textToWrite = DetailText;
                if (!string.IsNullOrEmpty(textToWrite) && textToWrite.Length > maxDisplayChars)
                {
                    textToWrite = textToWrite.Substring(startIndex: 0, length: maxDisplayChars) + "...";
                }

                var retVal = Level == 0
                                 ? $"{TimeStamp} ({TimeStamp.ToLocalTime().ToString("hh:mm:ss tt")} Local)  {Username}  {textToWrite}"
                                 : $"{TimeStamp.ToString("hh:mm:ss tt")} ({TimeStamp.ToLocalTime().ToString("hh:mm:ss tt")} Local)  {textToWrite}";
                retVal = retVal.Replace("    ", "  ");
                return retVal;
            }
        }


        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append("Id: ").Append(Id.ToString("0000"));
            sb.Append(", Version: ").Append(AppVersion);
            sb.Append($", Date/Time: {TimeStamp} ({TimeStamp.ToLocalTime()} Local)");
            sb.Append(", Level: ").Append(Level);
            sb.Append(", Description: ").Append(DetailText);
            sb.Append(", UserName: ").Append(Username);
            return sb.ToString();
        }


        public string ToOutputText()
        {
            var sb = new StringBuilder();
            sb.Append($"ID: {Id:0000}".WrapClean(80));
            sb.Append($"Version: {AppVersion}".WrapClean(80));
            sb.Append($"Date/Time: {TimeStamp} ({TimeStamp.ToLocalTime()} Local)".WrapClean(80));
            sb.Append($"Level: {Level}".WrapClean(80));
            sb.Append($"Description: {DetailText}".WrapClean(80));
            sb.Append($"UserName: {Username}".WrapClean(80));
            return sb.ToString();
        }

        protected override void OnDeserialize(AppSerializationState state)
        {
            foreach (var infoItem in state.Info)
            {
                var name = state.EffectiveItemName(infoItem.Name);
                switch (name)
                {
                case "ID":
                    Id = (int) infoItem.Value;
                    break;
                case "Date":
                    TimeStamp = (DateTime) infoItem.Value;
                    break;
                case "Level":
                    Level = (int) infoItem.Value;
                    break;
                case "Username":
                    Username = (string) infoItem.Value;
                    break;
                case "DetailText":
                    DetailText = (string) infoItem.Value;
                    break;
                case "ItemListText":
                    ItemText = (string) infoItem.Value;
                    break;
                case "AppVersion":
                    AppVersion = (string) infoItem.Value;
                    break;
                default:
                    state.AssertEntryNotHandled(name);
                    break;
                }
            }

            if (string.IsNullOrEmpty(AppVersion))
            {
                AppVersion = "< 3.1.1100";
            }
        }

        protected override void OnSerialize(AppSerializationState state)
        {
            var info = state.Info;
            info.AddValue("ID", Id);
            info.AddValue("Date", TimeStamp);
            info.AddValue("Level", Level);
            info.AddValue("Username", Username);
            info.AddValue("DetailText", DetailText);
            info.AddValue("ItemListText", ItemText);
            info.AddValue("AppVersion", AppVersion);
        }
    }
}