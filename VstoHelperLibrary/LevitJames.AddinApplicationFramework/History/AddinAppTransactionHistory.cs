// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using LevitJames.Core;
using LevitJames.MSOffice.MSWord;
using LevitJames.TextServices;

namespace LevitJames.AddinApplicationFramework
{
    /// <summary>
    ///     A keyed collection of HistoryRecord objects.
    /// </summary>
    
    [Serializable]
    public class AddinAppTransactionHistory : AppSerializableBase, IAddinAppProviderInternal
    {
        // FRONT MATTER

        private IEnumerable<AddinAppHistoryRecord> _items;

        public AddinAppTransactionHistory() { }

        protected AddinAppTransactionHistory(SerializationInfo info, StreamingContext context) : base(info, context) { }

        internal bool SessionHeaderWritten { get; set; }

        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        internal IAddinApplication App { get; private set; }


        public bool Enabled { get; set; }


        public IEnumerable<AddinAppHistoryRecord> Items => _items ?? Enumerable.Empty<AddinAppHistoryRecord>();


        public IEnumerable<AddinAppHistoryRecord> ItemsByDate => _items?.OrderBy(r => r.TimeStamp) ?? Enumerable.Empty<AddinAppHistoryRecord>();


        public bool Loaded => _items != null;


        private AddinAppDocument Document { get; set; }

        void IAddinAppProviderInternal.Initialize(IAddinApplicationInternal app, object data)
        {
            App = app;
            Document = (AddinAppDocument) data;
        }


        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        IAddinApplication IAddinAppProvider.App => App;


        /// <summary>Resets the input HistoryRecordCollection by clearing all but the last session record.</summary>
        public void Reset()
        {
            // Keep only the last session record
            var lastRecord = Items.LastOrDefault(hr => hr.Level == 0);
            if (lastRecord == null)
                return;

            Clear();
            AddRecord(lastRecord);
        }


        public void AddRecord(AddinAppHistoryRecord record)
        {
            App.Tracer.TraceInformation($"Transaction History: {record}");

            if (!Loaded)
                Load();

            EnsureSessionRecord();

            record.Id = GetNextId();
            record.TimeStamp = DateTime.UtcNow;
            record.Username = App.Environment.ResolvedUserName;
            record.AppVersion = App.Version.ToString();
            MarkAsDirty();
            ((List<AddinAppHistoryRecord>) _items).Add(record);

            App.Tracer.TraceInformation("Transaction History: Add out");
        }


        private void EnsureSessionRecord()
        {
            if (SessionHeaderWritten)
                return;

            SessionHeaderWritten = true;

            var sessionRecord = new AddinAppHistoryRecord(0, $"{App.Environment.ProductName} Session {App.Version}", string.Empty) {
                                                                                                                                       Id = GetNextId(),
                                                                                                                                       Username = App.Environment.ResolvedUserName,
                                                                                                                                       AppVersion = App.Version.ToString(),
                                                                                                                                       Level = 0,
                                                                                                                                       TimeStamp = DateTime.UtcNow
                                                                                                                                   };
            ((List<AddinAppHistoryRecord>) _items).Add(sessionRecord);
        }


        public void Load()
        {
            var history = Document.Store.GetHistory();
            _items = history != null ? new List<AddinAppHistoryRecord>(history._items) : new List<AddinAppHistoryRecord>();
        }


        public void Clear()
        {
            if (Loaded)
                ((List<AddinAppHistoryRecord>) _items).Clear();
        }

        private int GetNextId() => (!Items.Any() ? 0 : Items.Max(hr => hr.Id) + 1);


        protected override void OnDeserialize(AppSerializationState state)
        {
            _items = new List<AddinAppHistoryRecord>();

            App = state.Document.App;

            //Loop through the collection of values
            Document = state.Document;

            foreach (var infoItem in state.Info)
            {
                var name = state.EffectiveItemName(infoItem.Name);
                switch (name)
                {
                case "HistoryRecordList":
                    _items = (List<AddinAppHistoryRecord>) infoItem.Value;
                    break;
                case "Item0":
                    // It was serialized by old method
                    var items = new List<AddinAppHistoryRecord>();
                    Serializer.DeserializeCollection(state.Info, (AddinAppHistoryRecord hist) => items.Add(hist));
                    _items = items;
                    break;
                default:
                    if ((string.Compare(name, "Item1", StringComparison.Ordinal) >= 0 &&
                         string.Compare(name, "ItemZ", StringComparison.Ordinal) <= 0) ||
                        (name == "Count"))
                    {
                        continue;
                    }

                    state.AssertEntryNotHandled(name);
                    break;
                }
            }
        }


        protected override void OnSerialize(AppSerializationState state)
        {
            state.Info.AddValue("HistoryRecordList", _items);
        }


        public override string ToString()
        {
            if (!Loaded)
                Load();

            var entrySeparator = new string('*', 30);

            var sb = new StringBuilder();
            sb.AppendLine(entrySeparator);
            sb.AppendLine(entrySeparator);
            sb.Append($"{App.Environment.ProductName} History".WrapClean(80));
            sb.Append($"Document: {WordExtensions.WordApplication.WordBasicLJ().FileName}".WrapClean(80));
            sb.Append($"Saved on: {DateTime.UtcNow}".WrapClean(80));
            sb.AppendLine(entrySeparator);
            sb.AppendLine(entrySeparator);
            foreach (var hr in Items)
            {
                sb.Append(hr.ToOutputText());
                sb.Append(entrySeparator).AppendLine();
            }

            return sb.ToString();
        }
    }
}