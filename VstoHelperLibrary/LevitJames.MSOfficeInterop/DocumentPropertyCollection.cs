// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using LevitJames.Core;
using Microsoft.Office.Core;

namespace LevitJames.MSOffice
{
    /// <summary>
    ///     A collection of DocumentPropertyLJ objects, based on a provided collection of Word document properties.
    /// </summary>
    public sealed class DocumentPropertyCollection : ICollection<DocumentPropertyLJ>, IDisposable
    {
        private readonly dynamic _props;

#if (TRACK_DISPOSED)
        private readonly string _disposedSource;
#endif

        /// <summary>
        ///     Creates a new instance of the collection, wrapped around the provided collection of Word document properties
        /// </summary>
        /// <param name="props">The Word document properties collection to wrap</param>

        internal DocumentPropertyCollection(dynamic props)
        {
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
            _props = props;
        }


        /// <summary>
        ///     Retrieves an item in the collection by its index
        /// </summary>
        /// <param name="index">The integer index or string key of the document property to retrieve.</param>
        /// <returns>A DocumentPropertyLJ instance.</returns>
        
        public DocumentPropertyLJ this[object index] => new DocumentPropertyLJ(_props.Item[index]);


        void ICollection<DocumentPropertyLJ>.Add(DocumentPropertyLJ item)
        {
            throw new NotSupportedException();
        }


        /// <summary>
        ///     Removes all items in the collection
        /// </summary>
        
        public void Clear()
        {
            for (var i = _props.Count; i >= 1; i--)
            {
                var itm = _props.Item[i];
                itm.Delete();
                Marshal.ReleaseComObject(itm);
            }
        }


        /// <summary>
        ///     Returns true if the property collection contains the specified object
        /// </summary>
        /// <param name="item">The object to look for in the collection</param>
        /// <returns>true if the document property exists; false otherwise;</returns>
        
        public bool Contains(DocumentPropertyLJ item)
        {
            Check.NotNull(item, "item");

            for (var i = 1; i <= _props.Count; i++)
            {
                var itm = _props.Item[i];
                var result = itm == item.Instance;
                Marshal.ReleaseComObject(itm);
                if (result)
                    return true;
            }

            return false;
        }

        void ICollection<DocumentPropertyLJ>.CopyTo(DocumentPropertyLJ[] array, int arrayIndex)
        {
            for (var i = 1; i <= _props.Count; i++)
            {
                array[arrayIndex + i - 1] = new DocumentPropertyLJ(_props[i]);
            }
        }


        /// <summary>
        ///     Returns the number of items in the collection
        /// </summary>
        public int Count => (int) _props.Count;


        /// <summary>
        ///     Indicates if the collection is readonly. This value is always false.
        /// </summary>
        public bool IsReadOnly => false;


        /// <summary>
        ///     Removes the specified item from the collection
        /// </summary>
        /// <param name="item">The item to remove from the collection</param>
        /// <returns>true if the document property was successfully removed; false otherwise;</returns>
        public bool Remove(DocumentPropertyLJ item)
        {
            Check.NotNull(item, "item");
            item.Delete();
            return true;
        }


        /// <summary>
        ///     Returns an enumerator to iterate over the collection
        /// </summary>
        /// <returns>An enumerator of DocumentPropertyLJ items.</returns>
        public IEnumerator<DocumentPropertyLJ> GetEnumerator()
        {
            for (var i = 1; i <= _props.Count; i++)
            {
                var itm = _props[i];
                yield return new DocumentPropertyLJ(itm);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        ///     Adds a new property to the collection
        /// </summary>
        /// <param name="name">The name of the document property to add.</param>
        /// <param name="value">The string value of the property to add.</param>
        /// <param name="linkToContent">
        ///     Specifies whether the LinkToContent property is linked to the contents of the container
        ///     document. If this argument is True, the LinkSource argument is required; if it's False, the value argument is
        ///     required.
        /// </param>
        /// <param name="linkSource">
        ///     Ignored if LinkToContent is False. The source of the LinkSource property. The container
        ///     application determines what types of source linking you can use. For example, DDE links use the
        ///     "Server|Document!Item" syntax.
        /// </param>
        /// <returns>A new DocumentPropertyLJ instance.</returns>
        
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        [CLSCompliant(isCompliant: false)]
        public DocumentPropertyLJ Add(string name, string value, bool linkToContent = false, object linkSource = null)
        {
            return AddCore(_props, name, value, MsoDocProperties.msoPropertyTypeString, linkToContent, linkSource);
        }

        /// <summary>
        ///     Adds a new property to the collection
        /// </summary>
        /// <param name="name">The name of the document property to add.</param>
        /// <param name="value">The integer value of the property to add.</param>
        /// <param name="linkToContent">
        ///     Specifies whether the LinkToContent property is linked to the contents of the container
        ///     document. If this argument is True, the LinkSource argument is required; if it's False, the value argument is
        ///     required.
        /// </param>
        /// <param name="linkSource">
        ///     Ignored if LinkToContent is False. The source of the LinkSource property. The container
        ///     application determines what types of source linking you can use. For example, DDE links use the
        ///     "Server|Document!Item" syntax.
        /// </param>
        /// <returns>A new DocumentPropertyLJ instance.</returns>
        
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        [CLSCompliant(isCompliant: false)]
        public DocumentPropertyLJ Add(string name, int value, bool linkToContent = false, object linkSource = null)
        {
            return AddCore(_props, name, value, MsoDocProperties.msoPropertyTypeNumber, linkToContent, linkSource);
        }

        /// <summary>
        ///     Adds a new property to the collection
        /// </summary>
        /// <param name="name">The name of the document property to add.</param>
        /// <param name="value">The boolean value of the property to add.</param>
        /// <param name="linkToContent">
        ///     Specifies whether the LinkToContent property is linked to the contents of the container
        ///     document. If this argument is True, the LinkSource argument is required; if it's False, the value argument is
        ///     required.
        /// </param>
        /// <param name="linkSource">
        ///     Ignored if LinkToContent is False. The source of the LinkSource property. The container
        ///     application determines what types of source linking you can use. For example, DDE links use the
        ///     "Server|Document!Item" syntax.
        /// </param>
        /// <returns>A new DocumentPropertyLJ instance.</returns>
        
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        [CLSCompliant(isCompliant: false)]
        public DocumentPropertyLJ Add(string name, bool value, bool linkToContent = false, object linkSource = null)
        {
            return AddCore(_props, name, value, MsoDocProperties.msoPropertyTypeBoolean, linkToContent, linkSource);
        }

        /// <summary>
        ///     Adds a new property to the collection
        /// </summary>
        /// <param name="name">The name of the document property to add.</param>
        /// <param name="value">The DateTime value of the property to add.</param>
        /// <param name="linkToContent">
        ///     Specifies whether the LinkToContent property is linked to the contents of the container
        ///     document. If this argument is True, the LinkSource argument is required; if it's False, the value argument is
        ///     required.
        /// </param>
        /// <param name="linkSource">
        ///     Ignored if LinkToContent is False. The source of the LinkSource property. The container
        ///     application determines what types of source linking you can use. For example, DDE links use the
        ///     "Server|Document!Item" syntax.
        /// </param>
        /// <returns>A new DocumentPropertyLJ instance.</returns>
        
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        [CLSCompliant(isCompliant: false)]
        public DocumentPropertyLJ Add(string name, DateTime value, bool linkToContent = false, object linkSource = null)
        {
            return AddCore(_props, name, value.ToOADate(), MsoDocProperties.msoPropertyTypeDate, linkToContent, linkSource);
        }

        /// <summary>
        ///     Adds a new property to the collection
        /// </summary>
        /// <param name="name">The name of the document property to add.</param>
        /// <param name="value">The double value of the property to add.</param>
        /// <param name="linkToContent">
        ///     Specifies whether the LinkToContent property is linked to the contents of the container
        ///     document. If this argument is True, the LinkSource argument is required; if it's False, the value argument is
        ///     required.
        /// </param>
        /// <param name="linkSource">
        ///     Ignored if LinkToContent is False. The source of the LinkSource property. The container
        ///     application determines what types of source linking you can use. For example, DDE links use the
        ///     "Server|Document!Item" syntax.
        /// </param>
        /// <returns>A new DocumentPropertyLJ instance.</returns>
        
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed")]
        [CLSCompliant(isCompliant: false)]
        public DocumentPropertyLJ Add(string name, double value, bool linkToContent = false, object linkSource = null)
        {
            return AddCore(_props, name, value, MsoDocProperties.msoPropertyTypeFloat, linkToContent, linkSource);
        }

        private static DocumentPropertyLJ AddCore(dynamic props, string name, object value, MsoDocProperties type = 0, bool linkToContent = false, object linkSource = null)
        {
            if (TryGetItem(props, name, out DocumentPropertyLJ docProp))
            {
                if (docProp.Type == type)
                {
                    docProp.Value = value;
                    return docProp;
                }

                docProp.Delete();
                if (value == null)
                    return null;
            }

            if (value == null)
                return null;

            if (type != 0)
                type = GetDocPropertyType(ref value);

            var obj = props.Add(name, linkToContent, type, value, linkSource);

            return new DocumentPropertyLJ(obj);
        }


        internal static void AddOrUpdate(dynamic props, string name, object value)
        {
            var type = GetDocPropertyType(ref value);

            AddCore(props, name, value, type, false, null)?.Dispose();
        }


        private static MsoDocProperties GetDocPropertyType(ref object value)
        {
            if (value == null)
                return 0;

            if (value is DateTime)
            {
                value = Convert.ToDateTime(value).ToOADate();
                return MsoDocProperties.msoPropertyTypeDate;
            }

            if (value is string)
                return MsoDocProperties.msoPropertyTypeString;

            if (value is bool)
                return MsoDocProperties.msoPropertyTypeBoolean;

            bool isFixed;
            if (value.GetType().IsNumericType(out isFixed))
            {
                if (isFixed)
                    return MsoDocProperties.msoPropertyTypeNumber;
                return MsoDocProperties.msoPropertyTypeFloat;
            }

            throw new InvalidOperationException("value type is not supported");
        }

        /// <summary>
        ///     Returns true if the property collection contains the specified object
        /// </summary>
        /// <param name="name">The name of the property to look for in the collection</param>
        /// <returns>true if the document property exists; false otherwise;</returns>
        
        public bool Contains(string name)
        {
            for (var i = 1; i <= _props.Count; i++)
            {
                var itm = _props[i];
                if (itm.Name == name)
                    return true;
                Marshal.ReleaseComObject(itm);
            }

            return false;
        }


        /// <summary>
        ///     Returns the value of the DocumentProperty, or the supplied defaultValue if the property does not exist.
        /// </summary>
        /// <param name="name">The name of the property to retrieve</param>
        /// <param name="defaultValue">The default value to return if the property does not exist.</param>
        
        public object GetValue(string name, object defaultValue = null)
        {
            return GetValue(_props, name, defaultValue);
        }

        internal static object GetValue(dynamic props, string name, object defaultValue = null)
        {
            DocumentPropertyLJ itm;
            if (TryGetItem(props, name, out itm))
                return itm.Value;
            return defaultValue;
        }


        /// <summary>
        ///     Tries to get a document property for the name provided.
        /// </summary>
        /// <param name="name">The name of the document property to retrieve.</param>
        /// <param name="item">A DocumentPropertyLJ instance if exists; otherwise null;</param>
        /// <returns>true if the document property exists; false otherwise;</returns>
        public bool TryGetItem(string name, out DocumentPropertyLJ item)
        {
            return TryGetItem(_props, name, out item);
        }


        internal static bool TryGetItem(dynamic props, string name, out DocumentPropertyLJ item)
        {
            item = null;
            for (var i = 1; i <= props.Count; i++)
            {
                var itm = props[i];
                if (name.Equals(itm.Name, StringComparison.OrdinalIgnoreCase))
                {
                    item = new DocumentPropertyLJ(itm);
                    return true;
                }

                Marshal.ReleaseComObject(itm);
            }

            return false;
        }

        internal static bool Remove(dynamic props, string name)
        {
            for (var i = 1; i <= props.Count; i++)
            {
                var itm = props[i];
                if (!name.Equals(itm.Name, StringComparison.OrdinalIgnoreCase))
                    continue;

                itm.Delete();
                return true;
            }

            return false;
        }

        ~DocumentPropertyCollection()
        {
            Dispose(false);
        }
        public void Dispose(bool disposing)
        {
#if (TRACK_DISPOSED)
                LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            if (_props != null)
                Marshal.ReleaseComObject(_props);
 
        }
        /// <inheritdoc />
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}