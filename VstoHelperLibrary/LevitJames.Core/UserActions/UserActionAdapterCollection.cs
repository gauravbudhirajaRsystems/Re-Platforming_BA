// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    ///     A keyed collection containing UserActionAdapter instances
    /// </summary>
    /// <remarks>The key for each UserActionAdapter is the derived UserActionAdapter class type.</remarks>
    [EditorBrowsable(EditorBrowsableState.Advanced)]
    public sealed class UserActionAdapterCollection : KeyedCollection<Type, UserActionAdapter>
    {
        private readonly object _syncLock = new object();

        /// <summary>
        ///     Creates a new instance of UserActionAdapterCollection
        /// </summary>
        internal UserActionAdapterCollection() { }

        /// <summary>
        ///     Inserts an element into the <see cref="System.Collections.ObjectModel.KeyedCollection`2" /> at the specified
        ///     index.
        /// </summary>
        /// <param name="index">The zero-based index at which <paramref name="item" /> should be inserted.</param>
        /// <param name="item">The object to insert.</param>
        /// <exception cref="System.ArgumentOutOfRangeException">
        ///     <paramref name="index" /> is less than 0.-or-<paramref name="index" /> is greater than
        ///     <see cref="System.Collections.ObjectModel.Collection`1.Count" />.
        /// </exception>
        protected override void InsertItem(int index, UserActionAdapter item)
        {
            lock (_syncLock)
            {
                base.InsertItem(index, item);
            }
        }

        /// <summary>
        ///     Returns the class type of the item to use as the key for the item.
        /// </summary>
        /// <param name="item"></param>
        protected override Type GetKeyForItem([NotNull] UserActionAdapter item)
        {
            Check.NotNull(item, nameof(item));
            lock (_syncLock)
            {
                return item.GetType();
            }
        }

        /// <summary>
        ///     Removes the element at the specified index of the
        ///     <see cref="System.Collections.ObjectModel.KeyedCollection`2" />.
        /// </summary>
        /// <param name="index">The index of the element to remove.</param>
        protected override void RemoveItem(int index)
        {
            lock (_syncLock)
            {
                var item = this[index];
                base.RemoveItem(index);
                item.Clear();
            }
        }

        /// <summary>
        ///     Returns a UserActionAdapter of the Generic type provided, or null if the UserActionAdapter does not exist.
        /// </summary>
        /// <typeparam name="TUserActionAdapter"></typeparam>
        public TUserActionAdapter GetItem<TUserActionAdapter>() where TUserActionAdapter : UserActionAdapter
        {
            if (Dictionary == null)
                return default(TUserActionAdapter);

            lock (_syncLock)
            {
                UserActionAdapter adapter;
                Dictionary.TryGetValue(typeof(TUserActionAdapter), out adapter);
                return adapter as TUserActionAdapter;
            }
        }
    }
}