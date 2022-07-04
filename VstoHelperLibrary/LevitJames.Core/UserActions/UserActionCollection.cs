// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    /// <summary>
    ///     A collection containing UserAction classes
    /// </summary>
    /// <remarks>The key for each UserAction instance is the UserAction.Id.</remarks>
    [Serializable]
    public class UserActionCollection : KeyedCollection<string, UserAction>
    {
        private readonly object _syncLock = new object();

        /// <summary>
        ///     Adds a collection of UserActions to the collection in a single call.
        /// </summary>
        /// <param name="items">The UserActions to add</param>
        public void AddRange([ItemNotNull] params UserAction[] items)
        {
            Check.NotNull(items, nameof(items));
            lock (_syncLock)
            {
                foreach (var item in items)
                {
                    Add(item);
                }
            }
        }

        /// <summary>
        ///     Adds a collection of UserActions to the collection in a single call.
        /// </summary>
        /// <param name="items">The UserActions to add</param>
        public void AddRange([ItemNotNull] IEnumerable<UserAction> items)
        {
            Check.NotNull(items, nameof(items));
            lock (_syncLock)
            {
                foreach (var item in items)
                    Add(item);
            }
        }

        /// <summary>
        ///     Removes the range of UserActions from the collection.
        /// </summary>
        /// <param name="items"></param>
        public void RemoveRange([ItemNotNull] IEnumerable<UserAction> items)
        {
            if (items == null)
                return;

            lock (_syncLock)
            {
                foreach (var ua in items)
                    Remove(ua);
            }
        }

        /// <summary>
        ///     Clears all the UserActions in the collection.
        /// </summary>
        protected override void ClearItems()
        {
            RemoveRange(this.ToArray());
            base.ClearItems();
        }


        /// <summary>
        ///     Inserts an element into the <see cref="System.Collections.ObjectModel.KeyedCollection`2" /> at the specified
        ///     index.
        /// </summary>
        /// <param name="index">The zero-based index at which <paramref name="item" /> should be inserted.</param>
        /// <param name="item">The object to insert.</param>
        /// <exception cref="System.ArgumentOutOfRangeException">
        ///     The <paramref name="index" /> is less than 0.-or-
        ///     <paramref name="index" /> is greater than <see cref="System.Collections.ObjectModel.Collection`1.Count" />.
        /// </exception>
        protected override void InsertItem(int index, [NotNull] UserAction item)
        {
            Check.NotNull(item, nameof(item));
            lock (_syncLock)
            {
                base.InsertItem(index, item);
                UserActionManager.OnUserActionAdded(item);
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
                var userAction = this[index];
                base.RemoveItem(index);

                //Remove all the instances mapped to the UserAction
                UserActionManager.OnUserActionRemoved(userAction);
            }
        }

        /// <summary>
        ///     Returns the UserAction.Id to use as a key for the item.
        /// </summary>
        /// <param name="item"></param>
        protected override string GetKeyForItem([NotNull] UserAction item)
        {
            Check.NotNull(item, nameof(item));
            return item.Id;
        }

        /// <summary>
        ///     Gets the UserAction value associated with the specified key.
        /// </summary>
        /// <param name="id">The UserAction.Id of the UserAction to retrieve.</param>
        /// <param name="value">A UserAction instance; or null if not found.</param>
        [SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "1#")]
        public bool TryGetValue(string id, out UserAction value)
        {
            value = null;
            lock (_syncLock)
            {
                if (Dictionary == null)
                    return false;

                if (Dictionary.TryGetValue(id, out value))
                    return true;
            }

            return false;
        }
    }
}