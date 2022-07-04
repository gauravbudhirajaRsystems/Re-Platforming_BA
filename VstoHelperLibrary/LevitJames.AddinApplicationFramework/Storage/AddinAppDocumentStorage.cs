// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.Caching;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters;
using System.Runtime.Serialization.Formatters.Binary;
using JetBrains.Annotations;
using LevitJames.Core;
using LevitJames.MSOffice.MSWord;
using Microsoft.Office.Interop.Word;
using Version = System.Version;

namespace LevitJames.AddinApplicationFramework
{
    public class AddinAppDocumentStorage
    {
        protected internal const string ActiveNamedSessionVariableName = "ActiveNamedSession";
        protected internal string DocumentRecoveryVariableName = "DocumentRecovery";

        protected internal const string DisableWordEventsVariableName = "DisableWordEvents";

        protected internal const string DateFormatForVariables = "MM/dd/yyyy";

        protected internal AddinAppDocumentStorage([NotNull] AddinAppDocument document) : this(document.App.Environment)
        {
            Check.NotNull(document, nameof(document));
            Document = document;
            Instances = new InstanceStore(this);
        }

        protected internal AddinAppDocumentStorage([NotNull] AddinAppEnvironment environment)
        {
            Check.NotNull(environment, nameof(environment));
            StorageItemNamePrefix = $"{environment.SimpleCompanyName}.{environment.SimpleProductName}";
        }

        internal InstanceStore Instances { get; }

        protected internal virtual string StorageItemNamePrefix { get; }


        protected internal virtual bool IsAppDocument => Contains(MakeStorageName("History.First.SessionInfo"));

        protected internal virtual bool RecoveryRequired => GetBool(DocumentRecoveryVariableName);


        protected internal AddinAppDocument Document { get; }


        //[DebuggerDisplay("Count = {Count}")]
        public IEnumerable<string> Keys => GetKeys(new[] {StorageItemNamePrefix}, baseKeysOnly: false);
        public virtual IEnumerable<string> BaseKeys => GetKeys(new[] {StorageItemNamePrefix}, baseKeysOnly: true);
        public virtual IEnumerable<string> LegacyKeys => Enumerable.Empty<string>();


        public IEnumerable<StorageItem> Values => Keys.Select(k => new StorageItem(this, k));


        public IEnumerable<StorageItemCollection> UndoStore
        {
            get
            {
                if (UndoManager == null || !UndoManager.CanUndo)
                    return Enumerable.Empty<StorageItemCollection>();

                return UndoManager.GetUndoStackStore();
            }
        }

 
        public IEnumerable<StorageItemCollection> NestedTransactionStore
        {
            get
            {
                if (UndoManager == null || !UndoManager.CanUndo)
                    return Enumerable.Empty<StorageItemCollection>();

                return UndoManager.GetNestedTransactionStore();
            }
        }


        private UndoManager UndoManager { get; set; }

        internal bool CanUndo => UndoManager != null && UndoManager.CanUndo;

        internal string UndoDescription => UndoManager?.UndoDescription ?? string.Empty;

        internal bool Undoable => Document.Session.ActiveTransaction != null
                                  && Document.Session.ActiveTransaction.Undoable;


        internal bool ClearAllCalled { get; set; }

        //public event EventHandler<RestoreInstanceEventArgs> RestoreInstance;

        /// <summary>
        /// Returns a AddinAppTransactionHistory instance stored in the document.
        /// </summary>
        /// <param name="expression">A property expression locating where the property is stored in the model.</param>
        /// <returns>A AddinAppTransactionHistory instance</returns>
        public AddinAppTransactionHistory GetHistory(Expression<Func<AddinAppTransactionHistory>> expression = null)
            => Get(expression, undoable:false,key: MakeStorageName("History"));


        /// <summary>
        /// Expands the provided name to its full storage name in the form [CompanyName].[ProductName].name
        /// </summary>
        /// <param name="name">The short name of the variable</param>
        /// <returns></returns>
        /// <remarks>If [CompanyName].[ProductName] has already been appended it is not appended twice.</remarks>
        protected virtual string MakeStorageName(string name)
        {
            var prefix = StorageItemNamePrefix;
            if (name == null)
                return prefix;

            if (name.StartsWith(prefix))
                return name;

            return prefix + "." + name;
        }

        /// <summary>
        /// Returns a Session instance from the document.
        /// </summary>
        /// <param name="name">The session name to retrieve, if null the name defaults to 'Session'</param>
        /// <param name="last">true if this is the last session info instance;false otherwise</param>
        /// <returns></returns>
        public SessionInfo GetSessionInfo(string name = null, bool last = true)
        {
            var propName = $"History.{(last ? "Last" : "First")}{name ?? "Session"}Info";
            propName = MakeStorageName(propName);
            var value = Get(propName);
            if (!SessionInfo.TryParse(value, out SessionInfo sessionInfo))
                sessionInfo = new SessionInfo(Document.App.Version);

            return sessionInfo;
        }

        /// <summary>
        /// Stores a Session instance to the document.
        /// </summary>
        /// <param name="name">The session name to store, if null the name defaults to 'Session'</param>
        /// <param name="last">true if this is the last session info instance;false otherwise</param>
        /// <returns>A SessionInfo instance that was saved to the document </returns>
        public SessionInfo SetSessionInfo(string name = null, bool last = true)
        {
            var propName = $"History.{(last ? "Last." : "First.")}{name ?? "Session"}Info";
            propName = MakeStorageName(propName);
            var value = new SessionInfo(Document.App.Version);
            Set(propName, value.ToString(), nameIsShortForm: false, includeInCustomProperties: true);
            return value;
        }


        //Version Control


        /// <summary>Creates VersionControl object from value extracted out of Word.Document.Variable.</summary>
        /// <returns>Created VersionControl object.</returns>
        public AddInAppVersionControl GetVersionControl()
        {
            var version = Get("VersionControl");

            return version == null
                       ? UpdateVersionControl()
                       : // Return the current one.
                       new AddInAppVersionControl(version);
        }


        /// <summary>Updates the version control information stored in the Word.Document.</summary>
        public AddInAppVersionControl UpdateVersionControl()
        {
            var newVersionControl = Document.App.Environment.VersionControl.Clone();
            // Write to variable & custom document property
            Set("VersionControl", newVersionControl.ToString(), includeInCustomProperties: true);
            return newVersionControl;
        }

        /// <summary>
        /// Stores a flag in the document, that if not cleared, can be checked to determine if recovery is required. 
        /// </summary>
        protected virtual void SetRecoveryFlag() => SetBool(DocumentRecoveryVariableName, true);

        /// <summary>
        /// Clears the recovery flag from the document that was previously set with the SetRecoveryFlag call.
        /// </summary>
        protected internal virtual void ClearRecoveryFlag() => SetBool(DocumentRecoveryVariableName, false);

        /// <summary>
        /// Returns a custom property value from the document.
        /// </summary>
        /// <param name="name">The name of the property to retrieve</param>
        /// <param name="defaultValue">The default value of the property to return if the value does not exist.</param>
        /// <param name="nameIsShortForm">If true then the name of the property is expanded to include the company name and the product name in the form,
        /// [CompanyName].[ProductName].name</param>
        /// <returns>The custom property value or the default value passed in.</returns>
        protected virtual object GetCustomProperty(string name, object defaultValue = null, bool nameIsShortForm = true)
        {
            var propName = nameIsShortForm ? MakeStorageName(name) : name;
            return Document.WordDocument.GetCustomDocumentProperty(propName, defaultValue);
        }

        /// <summary>
        /// Returns a custom property string value from the document.
        /// </summary>
        /// <param name="name">The name of the property to retrieve</param>
        /// <param name="defaultValue">The default value of the property to return if the value does not exist.</param>
        /// <param name="nameIsShortForm">If true then the name of the property is expanded to include the company name and the product name in the form,
        /// [CompanyName].[ProductName].name</param>
        /// <returns>The custom property value or the default value passed in.</returns>
        protected virtual string GetCustomPropertyString(string name, string defaultValue = "", bool nameIsShortForm = true)
        {
            var propName = nameIsShortForm ? MakeStorageName(name) : name;
            return (string) Document.WordDocument.GetCustomDocumentProperty(propName, defaultValue);
        }

        /// <summary>
        /// Sets a custom property string value in the document.
        /// </summary>
        /// <param name="name">The name of the property to retrieve</param>
        /// <param name="value">The value to set</param>
        /// <param name="nameIsShortForm">If true then the name of the property is expanded to include the company name and the product name in the form,
        /// [CompanyName].[ProductName].name</param>
        /// <returns>The custom property value or the default value passed in.</returns>
        protected virtual void SetCustomProperty(string name, object value, bool nameIsShortForm = true)
        {
            var propName = nameIsShortForm ? MakeStorageName(name) : name;
            Document.WordDocument.SetCustomDocumentProperty(propName, value);
        }


        /// <summary>
        /// Gets a property stored in the document as a variable.
        /// </summary>
        /// <param name="name">The name of the property to retrieve</param>
        /// <param name="nameIsShortForm">If true then the name of the property is expanded to include the company name and the product name in the form,
        /// [CompanyName].[ProductName].name</param>
        /// <returns>The property value as a string.</returns>
        public string Get(string name, bool nameIsShortForm = true)
        {
            var variableName = nameIsShortForm ? MakeStorageName(name) : name;
            return ReadStringValue(variableName);
        }


        /// <summary>
        /// Gets a class instance that is optionally stored in the document.
        /// </summary>
        /// <typeparam name="TClass">The class type to restore. The type name of is used as the key for the instance.TClass should be a serializable type</typeparam>
        /// <param name="expression">A property expression that should start at the App.Document level. The property expression is used to clear the models value at the end of an edit session. It may also called when an undo or cancel operation occurs</param>

        /// <param name="persistToDocument">True to persist the instance in the document, false to not.</param>
        /// <param name="undoable">True, if the instance can be undone, or rolled-back. Generally this should be the same as the transaction type</param>
        /// <param name="priority">A numeric value that determines the order in which the class is serialized and retrieved during an undo or restore operation. This is required as some class types may depend on other types to be deserialized first.</param>
        /// <param name="key">An optional key to use to retrieve and store the class instance. Only used if there are multiple variables of the same class type stored in the document. If null the FullName of the TClass is used as the key.</param>
        /// <returns></returns>
        protected TClass Get<TClass>(Expression<Func<TClass>> expression, bool persistToDocument = true, bool undoable = true, int priority = 0, string key = null) where TClass : class
        {
            if (expression == null)
                return Get<TClass>(key);
 
            if (string.IsNullOrEmpty(key))
                key = typeof(TClass).FullName;

            var itm = Instances.Get(expression, key, persistToDocument, undoable, priority);
            if (undoable && Undoable)
                UndoManager.Add(itm);

            return itm.Instance as TClass;
        }

        /// <summary>
        /// Gets a class instance that is optionally stored in the document.
        /// </summary>
        /// <typeparam name="TClass">The class type to restore. The type name of is used as the key for the instance.TClass should be a serializable type</typeparam>
        /// <param name="key">An optional key to use to retrieve and store the class instance. Only used if there are multiple variables of the same class type stored in the document. If null the FullName of the TClass is used as the key.</param>
        /// <returns></returns>
        protected TClass Get<TClass>(string key = null) where TClass : class
        {
            if (string.IsNullOrEmpty(key))
                key = typeof(TClass).FullName;

            try
            {
                var instance = ReadValue(key);
                return (TClass) instance;
            }
            finally
            {
                Document.FillProxy = false;
            }
        }

        /// <summary>
        /// Stores a variable name value pair in the document variables.
        /// </summary>
        /// <param name="name">The name of the variable to store.</param>
        /// <param name="value">The string value to store.</param>
        /// <param name="includeInCustomProperties">True to also duplicate the property name/value in the document properties.</param>
        /// <param name="nameIsShortForm">If true then the name of the property is expanded to include the companyname and the product name in the form,
        /// [CompanyName].[ProductName].name</param>
        public void Set(string name, string value, bool includeInCustomProperties = false, bool nameIsShortForm = true)
        {
            Document.App.Tracer.TraceInformation($"Store.Set {name} = {value}");

            var variableName = nameIsShortForm ? MakeStorageName(name) : name;
            WriteValue(variableName, value);
            if (includeInCustomProperties)
                SetCustomProperty(variableName, value, nameIsShortForm: true);
        }

        /// <summary>
        /// Stores a variable name value pair in the document variables. The type name of TItem is used as the key. 
        /// </summary>
        /// <typeparam name="TItem"></typeparam>
        /// <param name="value"></param>
        protected internal void Set<TItem>(TItem value) where TItem : class
        {
            Set(typeof(TItem).FullName, value);
        }

        /// <summary>
        /// Stores the class instance of type TItem in the document.
        /// </summary>
        /// <typeparam name="TItem"></typeparam>
        /// <param name="key">The key used to store the value in the document. This should be in the short form (without the company name and product name)</param>
        /// <param name="value">The class instance to store</param>
        /// <remarks>In order to store and retrieve a class instance, the class should be marked as serializable and implement ISerializable, or inherit from AddinAppSerializableBase. </remarks>
        protected internal void Set<TItem>(string key, TItem value) where TItem : class
        {
            if (value is string)
            {
                //Special case in-case TItem get inferred as a string.
                Set(key, value as string, includeInCustomProperties: false, nameIsShortForm: true);
                return;
            }
 
            WriteValue(key, value);
        }

        /// <summary>
        /// Returns a boolean flag value from the document variables.
        /// </summary>
        /// <param name="name">The name of the variable to retrieve.</param>
        /// <param name="defaultValue">The default value of the variable</param>
        /// <param name="nameIsShortForm">If true then the name of the property is expanded to include the companyname and the product name in the form,
        /// [CompanyName].[ProductName].name</param>
        /// <param name="delete">True to delete the variable after retrieving it.</param>
        /// <returns></returns>
        public bool GetBool(string name, bool defaultValue = false, bool nameIsShortForm = false, bool delete = false)
        {
            var variableName = nameIsShortForm ? name : MakeStorageName(name);

            var value = (string) ReadValue(variableName, defaultValue.ToString());

            if (delete)
                Document.WordDocument.Variables.DeleteLJ(variableName);

            bool.TryParse(value, out bool boolValue);
            return boolValue;
        }

        /// <summary>
        /// Sets a boolean flag value from the document variables.
        /// </summary>
        /// <param name="name">The name of the variable to set.</param>
        /// <param name="value">The value to set</param>
        /// <param name="keepSavedState">True to keep the Document.Saved state after setting the value.</param>
        /// <param name="includeInCustomProperties">True to duplicate the value in the Document.Properties collection.</param>
        /// <param name="removeIfFalse">If the value provided is false, then do not store the value and remove it if it exists.</param>
        public void SetBool(string name, bool value, bool keepSavedState = true, bool includeInCustomProperties = false, bool removeIfFalse = true)
        {
            var origDocSavedState = Document?.WordDocument?.Saved;

            if (origDocSavedState == null)
                return;

            try
            {
                var variableValue = value.ToString();
                var variableName = MakeStorageName(name);

                Document.App.Tracer.TraceInformation($"{name} {(value ? "Removing" : "Adding")}");
                if (value == false && removeIfFalse)
                {
                    Document.WordDocument.Variables.DeleteLJ(variableName);
                    if (includeInCustomProperties)
                        Document.WordDocument.RemoveCustomDocumentProperty(variableName);
                    return;
                }

                WriteValue(variableName, variableValue);
                if (includeInCustomProperties)
                    SetCustomProperty(variableName, value, nameIsShortForm: true);
            }
            finally
            {
                if (keepSavedState)
                    Document.WordDocument.Saved = origDocSavedState == true;
            }
        }

        /// <summary>
        /// Returns all the keys in teh document, that start with the provided prefixes.
        /// </summary>
        /// <param name="prefixes">An array of prefixes. Typically there will be only one prefix. But for compatibility other prefixes may exist in the document.</param>
        /// <param name="baseKeysOnly">True to only return the base bases (The first key, when an item split into multiple keys)</param>
        /// <returns></returns>
        protected virtual IEnumerable<string> GetKeys(string[] prefixes, bool baseKeysOnly)
        {
            foreach (var keyPrefix in prefixes)
            {
                foreach (Variable v in Document.WordDocument.Variables)
                {
                    var name = v.Name;
                    if (name.StartsWith(keyPrefix) && (!baseKeysOnly || IsBaseItemName(name)))
                        yield return name;
                }
            }
        }


        /// <summary>
        /// internal call to begin restoring instances.
        /// </summary>
        internal void OnRestoreInstances()
        {
            var itemsToRestore = Instances.GetInstancesToRestore();
            if (itemsToRestore.Count > 0)
                OnRestoreInstances(new RestoreInstancesEventArgs { Items = itemsToRestore });
        }

        /// <summary>
        /// Called when storage instances are about to be restored
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnRestoreInstances(RestoreInstancesEventArgs e)
        {
            foreach (var item in e.Items)
                OnRestoreInstance(new RestoreInstanceEventArgs { Item = item });
        }

        /// <summary>
        /// Called when a storage instance is about to be restored
        /// </summary>
        /// <param name="e"></param>
        protected internal virtual void OnRestoreInstance(RestoreInstanceEventArgs e) => Instances.OnRestoreInstance(e);

        internal bool UndoInternal() => UndoManager != null && UndoManager.Undo();

        /// <summary>
        /// Called when a transaction has completed, but is still the ActiveTransaction
        /// </summary>
        protected internal virtual void OnCompleteTransaction()
        {
            UndoManager?.CompleteTransaction();
        }


        internal void BeginTransaction()
        {
            UndoManager?.BeginTransaction();
        }

        /// <summary>
        /// Returns if a variable key exists in the document. The key name is in long form. [CompanyName].ProductName
        /// </summary>
        /// <param name="key"></param>
        /// <param name="nameIsShortForm">If true then the name of the property is expanded to include the company name and the product name in the form,
        /// [CompanyName].[ProductName].name</param>
        protected virtual bool Contains(string key,bool nameIsShortForm = false) => Document.WordDocument.Variables.ExistsLJ(key);

        /// <summary>
        /// Creates the stream instance used to serialize class items in the document.
        /// </summary>
        /// <param name="key">The key to the item being serialized.</param>
        /// <returns></returns>
        protected virtual Stream CreateStream(string key) => new AssertableMemoryStream("Key=" + key);


        /// <summary>
        ///     Returns false if the name of the variable ends in a number 00001 etc; otherwise returns true;
        /// </summary>
        /// <param name="name">The name of the variable to check.</param>
        
        protected internal virtual bool IsBaseItemName([NotNull] string name)
        {
            var idx = name.LastIndexOf('.');
            if (idx != -1 && (name.Length - idx) == 5) // look to see if the name ends with .0001 suffix.
            {
                if (int.TryParse(name.Substring(idx + 1), out int _))
                    return false; // it ends in a numeric number so skip
            }

            return true;
        }
        /// <summary>
        /// Resolves an Instance Key (the key used for the instance) into a Store Key (the key saved in the document).
        /// </summary>
        /// <returns>The default implementation does nothing, this is, it simply returns the key passed in.</returns>
        /// <remarks>This override can be useful for compatibility when a key stored in the document changes between versions. </remarks>
        protected virtual string InstanceKeyToStoreKey(string key) => key;

        /// <summary>
        /// Resolves a Store key (the key saved in the document) into an Instance Key (the used for the instance).
        /// </summary>
        /// <returns>The default implementation does nothing, this is, it simply returns the key passed in.</returns>
        /// <remarks>This override can be useful for compatibility when a key stored in the document changes between versions. </remarks>
        protected virtual string StoreKeyToInstanceKey(string key) => key;

        /// <summary>
        /// Writes the string value to a document variable. If the string is very large it is split into segments ending with a numeric suffix in the form key.[0000] 
        /// </summary>
        /// <param name="key">The name of the variable used to store the value.</param>
        /// <param name="value">The string value to store.</param>
        /// <remarks>Every class that is serialized will call this method when it stores itself in the document. So it can be useful for debugging when a class gets stored in the document. </remarks>
        protected virtual void WriteStringValue([NotNull] string key, string value)
        {
            if (value == null)
                value = string.Empty;

            var remainder = value.Length;
            var count = -1;
            var pos = 0;
            var variables = Document.WordDocument.Variables;

            var resolvedKey = InstanceKeyToStoreKey(key);

            do
            {
                count++;
                var workName = resolvedKey + (count == 0 ? null : "." + count.ToString("0000", CultureInfo.InvariantCulture));

                if (remainder > 0)
                {
                    var len = Math.Min(remainder, val2: 32000);
                    variables[workName].Value = value.Substring(pos, len);
                    pos += len;
                    remainder -= len;
                    continue;
                }

                if (!variables.ExistsLJ(workName))
                    break;

                variables.DeleteLJ(workName);
            } while (count < 1000);
        }

   
        /// <summary>
        /// Gets a string value stored in a document variable with the supplied key.
        /// </summary>
        /// <param name="key">The key name of the variable to retrieve from the document</param>
        /// <returns></returns>
        protected internal virtual string ReadStringValue([NotNull] string key)
        {
            var resolvedKey = InstanceKeyToStoreKey(key);

            string retVal = null;
            var count = 0;
            var variables = Document.WordDocument.Variables;

            do
            {
                var workName = resolvedKey + (count == 0 ? null : "." + count.ToString("0000", CultureInfo.InvariantCulture));
                if (!variables.TryGetValueLJ(workName, out string workVal))
                    break;

                retVal += workVal;
                count++;
            } while (count > 0);

            return retVal;
        }

        /// <summary>
        /// Writes a value to the document, optionally serializing the value and updating the store.
        /// </summary>
        /// <param name="key">The key used to store the value in the document.</param>
        /// <param name="value">The value to store, this can be a string a stream or a serializable class instance.</param>
        protected internal void WriteValue(string key, object value)
        {
            WriteValue(key, value, updateStore: true);
        }
 
        internal void WriteValue(string key, object value, bool updateStore)
        {
            if (value == null || value is string)
            {
                WriteStringValue(key, (string) value);
                return;
            }
            
            var stream = value as Stream;
            var disposeStream = false;
            object instance = null;
            if (stream == null)
                instance = value;

            Stream streamToStore = null;
            if (stream == null)
            {
                stream = InstanceToStream(instance, key);
                streamToStore = stream;
                disposeStream = !updateStore;
            }
            else if (updateStore)
            {
                streamToStore = new MemoryStream();
                stream.CopyTo(streamToStore);
                streamToStore.Position = 0;
            }


            try
            {
                var base64String = Convert.ToBase64String(StreamToByteArray(stream));
                WriteStringValue(key, base64String);

                if (updateStore)
                    Instances.TryUpdateStoreItem(key, streamToStore, instance);
                    
            }
            finally
            {
                if (disposeStream)
                    streamToStore?.Dispose();
            }
        }

        /// <summary>
        /// Returns a stream instance for the provided key 
        /// </summary>
        /// <param name="key">The variable key to return the stream for.</param>
        /// <returns></returns>
        protected internal Stream ReadStream(string key)
        {
            var base64String = ReadStringValue(key);
            if (string.IsNullOrEmpty(base64String))
                return null;

            var bytes = Convert.FromBase64String(base64String);
            var stream = CreateStream(key);
            stream.Write(bytes, 0, bytes.Length);
            stream.Flush();
            stream.Position = 0;
            return stream;
        }

        /// <summary>
        /// Returns an value for the provided key. The value returned may be a simple string value or a complex class instance depending on the value stored.
        /// </summary>
        /// <param name="key">The variable key to return the stream for.</param>
        /// <param name="defaultValue">A default object to return if the value does not exist.</param>
        /// <returns></returns>
        protected internal object ReadValue(string key, object defaultValue = null)
        {
            const string objectStreamHeader = "AAEAAAD/////";

            var value = ReadStringValue(key);
            if (string.IsNullOrEmpty(value))
                return defaultValue;

            if (!value.StartsWith(objectStreamHeader))
                return value;

            using (var stream = ReadStream(key))
            {
                return InstanceFromStream(stream, key);
            }
        }

        /// <summary>
        /// Serializes the provided instance to a stream.
        /// </summary>
        /// <param name="value">The value to serialize. If the stream is null, null is returned.</param>
        /// <param name="key">The key name of the instance.</param>
        /// <param name="stream">An optional stream to use to store the serialized instance. If provided this stream is returned.</param>
        /// <returns></returns>
        protected internal Stream InstanceToStream(object value, string key, Stream stream = null)
        {
            if (value == null)
                return null;

            // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
            if (value is string)
            {
                if (stream == null)
                    stream = CreateStream(key);
                var bytes = Convert.FromBase64String((string) value);
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush();
                stream.Position = 0;
                return stream;
            }

            if (stream == null)
                stream = CreateStream(value.GetType().Name);

            CreateFormatter().Serialize(stream, value);
            return stream;
        }

        /// <summary>
        /// Returns a class instance from the provided stream.
        /// </summary>
        /// <param name="stream">The stream to de-serialize</param>
        /// <param name="key">The key of the instance to restore</param>
        /// <returns></returns>
        protected internal object InstanceFromStream(Stream stream, string key)
        {
            if (stream == null)
                return null;
            stream.Position = 0;
            return CreateFormatter().Deserialize(stream);
        }


        private byte[] StreamToByteArray(Stream input)
        {
            if (input == null || input.Length == 0)
                return Array.Empty<byte>();

            input.Position = 0;
            if (input is MemoryStream ms)
                return ms.ToArray();

            using (ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }
        /// <summary>
        /// Returns an instance that implements IFormatter. This is used to serialize and de-serialize instances to and from the document.
        /// The default implementation returns a BinaryFormatter instance. Note that if a custom instance is returned that the StreamingContext should pass a AddinAppDocument instance to the additional member. 
        /// </summary>
        /// <returns></returns>
        protected internal virtual IFormatter CreateFormatter() =>
            new BinaryFormatter(selector: null, context: new StreamingContext(StreamingContextStates.All, Document))
            { AssemblyFormat = FormatterAssemblyStyle.Simple };

        /// <summary>
        /// Compares two streams that were serialized by the store, and returns there differences as key/value pairs. 
        /// </summary>
        /// <param name="stream1">A stream instance to compare with stream2</param>
        /// <param name="stream2">A stream instance to compare with stream1</param>
        /// <returns>A collection of SerializationStreamComparerSet containing the differences between the tow streams.</returns>
        /// <remarks>Since comparing streams this way is slow, this is only done in debug builds.</remarks>
        protected internal virtual IEnumerable<SerializationStreamComparerSet> CompareStreams(Stream stream1, Stream stream2)
        {
            stream1.Position = 0;
            stream2.Position = 0;
            var comparer = new SerializationStreamComparer();

            SerializationStreamComparerDelegate callback = CompareStreamsCallback;

            return comparer.Compare(stream1, stream2, CreateFormatter(), callback);
        }

        private bool CompareStreamsCallback(string key, object value1, object value2, Version storedVersion)
        {
            if (key == Serializer.SerializeVersionName)
                return true;
            if (value1 == SerializationStreamComparer.MissingValue || value2 == SerializationStreamComparer.MissingValue)
            {
                if (AppSerializationState.EffectiveItemNameInternal(key, storedVersion) == null)
                    return true;
            }

            return Equals(value1, value2);
        }


        internal void OnEnterSession()
        {
            //SetSessionInfo();
            SetRecoveryFlag();

            if (UndoManager != null)
                return;

            UndoManager = new UndoManager(Document);
            var cachedInstances = (IEnumerable<StoreItem>) Document.AppInternal.TransactionManager.Cache.Get(GetHashCode().ToString());
            if (cachedInstances != null)
                Instances.Restore(cachedInstances);

            UndoManager.OnEnterSession();
        }
     
        internal void OnExitSession()
        {
            if (!Document.Session.InSession || Document.Session.InNamedSession())
                return;

            ClearRecoveryFlag();

            var instanceItems = Instances.Items.ToList();
            Instances.Clear();

            //we only call OnExitSession if we are not in a named session
            UndoManager.OnExitSession();

            //Cache the objects.
            if (Document.AppInternal.Environment.DocumentStorageCacheTimeout > 0)
            {
                var policy = new CacheItemPolicy {AbsoluteExpiration = DateTimeOffset.Now.AddSeconds(Document.AppInternal.Environment.DocumentStorageCacheTimeout)};

                var cacheItem = new CacheItem(GetHashCode().ToString(), instanceItems);
                Document.AppInternal.TransactionManager.Cache.Add(cacheItem, policy);
            }

            UndoManager = null;
        }

        /// <summary>
        /// Clears all the storage used for undoing transactions
        /// </summary>
        public void ClearUndo()
        {
            UndoManager?.Clear();
        }

        /// <summary>
        /// Clears the variables in the document using the provided keys.
        /// </summary>
        /// <param name="keysToClear">A collection of keys to clear.</param>
        /// <remarks>The Word undo stack is also cleared.</remarks>
        public void ClearAll(IEnumerable<string> keysToClear)
        {
            if (keysToClear == null || !keysToClear.Any())
                return;

            var wordDoc = Document.WordDocument;
            var customProps = wordDoc.CustomDocumentPropertiesLJ();

            var currentUndoItems = UndoManager?.CurrentData();
   
            foreach (var key in keysToClear)
            {
                if (customProps.TryGetItem(key, out var prop))
                    prop.Delete();
                wordDoc.Variables.DeleteLJ(key);

                //Reset (and null) the instance
                var instanceKey = StoreKeyToInstanceKey(key);
                Instances.RemoveItem(instanceKey);
                //Reset the instance if it is in the current undo stack
                var undoItem = currentUndoItems?.FirstOrDefault(i => i.Key == instanceKey);
                if (undoItem != null)
                    currentUndoItems.Remove(undoItem);

            }
 
            wordDoc.UndoClear();

            ClearAllCalled = true;
        }

          
        /// <summary>
        /// Clears all the variables in the document.
        /// </summary>

        public void ClearAll()
        {
            var keysToClear = Keys.Union(LegacyKeys);
            ClearAll(keysToClear);
        }


    }
}