// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using LevitJames.Core;

namespace LevitJames.AddinApplicationFramework
{
    public sealed class AppSerializationState
    {
        private AppSerializationState(SerializationInfo info, StreamingContext context) : this((AddinAppDocument) context.Context, info.ObjectType)
        {
            Document = (AddinAppDocument) context.Context;
            Info = info;
            Context = context;
        }

        internal AppSerializationState(AddinAppDocument document, Type objectType)
        {
            Document = document;
            AppVersion = document.App.Version;
            ClassVersion = objectType.GetCustomAttribute<AddinAppClassSerializeVersionAttribute>(inherit: true)?.Version
                           ?? AppVersion;
        }

        public AddinAppDocument Document { get; private set; }

        public SerializationInfo Info { get; }
        public Version ClassVersion { get; }
        public Version StoredVersion { get; private set; }
        public Version AppVersion { get; }

        public StreamingContext Context { get; }

        public static AppSerializationState OnSerialize(SerializationInfo info, StreamingContext context)
        {
            Check.NotNull(info, nameof(info));
            var state = new AppSerializationState(info, context);
            Serializer.SetVersionInfo(info, state.ClassVersion);
            return state;
        }

        public static AppSerializationState OnDeserialize(object instance, SerializationInfo info, StreamingContext context)
        {
            Check.NotNull(instance, nameof(instance));
            Check.NotNull(info, nameof(info));

            var s = new AppSerializationState(info, context);
            if (s.Document != null)
                (instance as IAddinAppDocumentProvider)?.SetDocument(s.Document);

            s.StoredVersion = Serializer.ReadVersion(info, s.ClassVersion.ToString());

            return s;
        }

        public void AssertEntryNotHandled(string itemName)
        {
            if (itemName == null || Serializer.SerializeVersionName == itemName)
                return;

            if (Document.AppInternal.UserSettings.Internal.AssertEntryNotHandled)
                Debug.Assert(StoredVersion > AppVersion, $"Class {Info.ObjectType.Name} has an unhandled serialized property '{itemName}'");
            else
                Debug.WriteLine($"Class {Info.ObjectType.Name} has an unhandled serialized property '{itemName}'");
        }


        public string EffectiveItemName(string serializedName) => EffectiveItemNameInternal(serializedName, AppVersion);

        public static string EffectiveItemNameInternal(string serializedName, Version appVersion)
        {
            if (serializedName == Serializer.SerializeVersionName)
                return null;

            var workName = serializedName;
            while (workName.Length > 3 && "<>=".Contains(workName[index: 0].ToString()))
            {
                var compOperator = workName.Substring(startIndex: 0, length: workName[index: 1] == '=' ? 2 : 1);
                var delimiterPosition = workName.IndexOf(value: ':');
                var versionLength = delimiterPosition - compOperator.Length;
                if (versionLength < 1)
                {
                    throw new ArgumentException(@"Invalid sourceText " + serializedName, nameof(serializedName));
                }

                var versionString = workName.Substring(compOperator.Length, versionLength);
                var compVersion = new Version(versionString);

                var match = false;
                switch (compOperator)
                {
                case "<":
                    match = appVersion < compVersion;
                    break;
                case "<=":
                    match = appVersion <= compVersion;
                    break;
                case "=":
                    match = appVersion == compVersion;
                    break;
                case ">=":
                    match = appVersion >= compVersion;
                    break;
                case ">":
                    match = appVersion > compVersion;
                    break;
                }

                if (match == false)
                {
                    //*************************************************
                    // Filter condition failed - pageItem should be ignored
                    return null;
                    //*************************************************
                }

                if (delimiterPosition + 1 >= workName.Length)
                {
                    throw new ArgumentException(@"Invalid sourceText " + serializedName, nameof(serializedName));
                }

                // Left trim everything before ":"
                workName = workName.Substring(delimiterPosition + 1);
            }

            return workName;
        }


        /// <summary>
        ///     Serializes the supplied instance into SerializationInfo as a byte array
        /// </summary>
        /// <param name="name">The name for the stored item</param>
        /// <param name="value">A class instance to serialize into SerializationInfo</param>
        public void AddBinaryValue(string name, object value)
        {
            if (value == null)
                return;

            using (var stream = new MemoryStream())
            {
                Document.Store.CreateFormatter().Serialize(stream, value);
                stream.Position = 0;
                Info.AddValue(name, stream.ToArray());
            }
        }


        /// <summary>
        ///     Deserializes an instance that was previously serialized using AddBinaryValue.
        /// </summary>
        /// <param name="name">The name of the stored item</param>
        public object GetBinaryValue(string name)
        {
            using (var stream = new MemoryStream((byte[]) Info.GetValue(name, typeof(byte[]))))
                return Document.Store.CreateFormatter().Deserialize(stream);
        }

        /// <summary>
        ///     Deserializes an instance that was previously serialized using AddBinaryValue.
        /// </summary>
        /// <param name="item">A SerializationEntry for the item to retrieve.</param>
        
        public object GetBinaryValue(SerializationEntry item)
        {
            using (var stream = new MemoryStream((byte[]) item.Value))
                return Document.Store.CreateFormatter().Deserialize(stream);
        }
    }
}