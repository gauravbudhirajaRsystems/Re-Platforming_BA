// © Copyright 2018 Levit & James, Inc.

using System;
using System.Runtime.Serialization;
using System.Security.Permissions;

namespace LevitJames.AddinApplicationFramework
{
    public abstract class AddinAppSerializableBase<TAppDocument> : AppSerializableBase, IAddinAppDocumentProvider where TAppDocument : AddinAppDocument, new()
    {
        protected AddinAppSerializableBase() { }

        protected AddinAppSerializableBase(SerializationInfo info, StreamingContext context) : base(info, context) { }

        //public TAppDocument Document { get; private set; }
        protected TAppDocument Document { get; private set; }

        void IAddinAppDocumentProvider.SetDocument(AddinAppDocument document)
        {
            SetAddinAppDocument(document);
        }

        protected virtual void SetAddinAppDocument(AddinAppDocument document)
        {
            Document = (TAppDocument) document;
        }
    }

    [Serializable]
    public abstract class AppSerializableBase : ISerializable, IAddinAppDirty
    {
        private int _dirtyCookie;

        protected internal AppSerializableBase() { }

        protected AppSerializableBase(SerializationInfo info, StreamingContext context)
        {
            // ReSharper disable once VirtualMemberCallInConstructor
            OnDeserialize(AppSerializationState.OnDeserialize(this, info, context));
        }

        protected virtual bool AutoDirty => true;


        public bool Dirty => ((IAddinAppDirty) this).DirtyCookie != 0;

        int IAddinAppDirty.DirtyCookie
        {
            get => AutoDirty ? -1 : _dirtyCookie;
            set
            {
                if (!AutoDirty) _dirtyCookie = value;
            }
        }

        [SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.SerializationFormatter)]
        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
            OnSerialize(AppSerializationState.OnSerialize(info, context));
        }

        protected virtual void OnDeserialize(AppSerializationState state) { }

        protected virtual void OnSerialize(AppSerializationState state) { }

        public void MarkAsDirty() => MarkAsDirtyGenerator.MarkAsDirty(this);
    }
}