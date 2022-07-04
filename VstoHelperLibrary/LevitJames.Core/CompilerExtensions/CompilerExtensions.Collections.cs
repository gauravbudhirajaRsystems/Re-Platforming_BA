// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using JetBrains.Annotations;

namespace LevitJames.Core
{
    //[DebuggerStepThrough]
    public static partial class CompilerExtensions
    {
        /// <summary>
        ///     Populates the array with the value provided
        /// </summary>
        /// <typeparam name="T">The Type of the Array</typeparam>
        /// <param name="source">The array to populate</param>
        /// <param name="value">The value to populate all the items in the array with.</param>
        public static T[] Populate<T>(this T[] source, T value)
        {
	        if (source == null)
		        return null;

            for (var i = 0; i < source.Length; i++)
			        source[i] = value;
            return source;
        }

        //// <summary>
        ////     Converts a character array to string array.
        //// </summary>
        //public static string[] ToStringArray(this char[] source) {
        //    var stringArray = new string[source.Length];
        //    for (var i = 0; i < stringArray.Length; i++)
        //        stringArray[i] = source[i].ToString();
        //    return stringArray;
        //}


        /// <summary>Determines the index of a specific item in the <see cref="IList{T1}" />.</summary>
        /// <param name="source">The list to find the item in</param>
        /// <param name="value">The object to locate in the <see cref="IList{T1}" />. </param>
        /// <param name="comparer"></param>
        /// <param name="startIndex">The starting index of the item to find</param>
        /// <param name="endIndex">The ending index of the item to find. default is -1 meaning the Count of the collection is used</param>
        /// <returns>The index of <paramref name="value" /> if found in the list; otherwise, -1.</returns>
        [DebuggerStepThrough]
        public static int IndexOf<T1, T2>(this IList<T1> source, T2 value, Func<T1, T2, bool> comparer, int startIndex, int endIndex = -1)
        {
            if (source == null)
                return -1;
            if (endIndex == -1)
                endIndex = source.Count;

            for (var i = startIndex; i < endIndex; i++)
                if (comparer(source[i], value))
                    return i;

            return -1;
        }

        /// <summary>Determines the previous index of a specific item in the <see cref="IList{T1}" />.</summary>
        /// <param name="source">The list to find the item in</param>
        /// <param name="value">The object to locate in the <see cref="IList{T1}" />. </param>
        /// <param name="startIndex">The starting index of the item to find. The default is -1 meaning its starts at the end of the collection.</param>
        /// <param name="endIndex">The ending index of the item to find. default is 0</param>
        /// <returns>The index of <paramref name="value" /> if found in the list; otherwise, -1.</returns>
        [DebuggerStepThrough]
        public static int IndexOfReverse<T>(this IList<T> source, T value, int startIndex = -1, int endIndex = 0)
        {
            if (source == null)
                return -1;

            if (startIndex == -1)
                startIndex = source.Count;

            for (var i = startIndex; i >= endIndex; i--)
                if (Equals(source[i], value))
                    return i;

            return -1;
        }

        /// <summary>Determines the previous index of a specific item in the <see cref="IList{T1}" />.</summary>
        /// <param name="source">The list to find the item in</param>
        /// <param name="comparer"></param>
        /// <param name="value">The object to locate in the <see cref="IList{T1}" />. </param>
        /// <param name="startIndex">The starting index of the item to find. The default is -1 meaning its starts at the end of the collection.</param>
        /// <param name="endIndex">The ending index of the item to find. default is 0</param>
        /// <returns>The index of <paramref name="value" /> if found in the list; otherwise, -1.</returns>
        [DebuggerStepThrough]
        public static int IndexOfReverse<T1, T2>(this IList<T1> source, T2 value, Func<T1, T2, bool> comparer, int startIndex = -1, int endIndex = 0)
        {
            if (source == null)
                return -1;

            if (startIndex == -1)
                startIndex = source.Count;

            for (var i = startIndex; i >= endIndex; i--)
                if (comparer(source[i], value))
                    return i;

            return -1;
        }

        /// <summary>Loops through a collection of items calls IDisposable.Dispose on each item before clearing the collection</summary>
        /// <typeparam name="T">A type that implements IDisposible</typeparam>
        /// <param name="source">The collection of items to dispose and clear. Can be null.</param>
        public static void ClearAndDispose<T>(this ICollection<T> source) where T : IDisposable
        {
            if (source == null)
                return;

            foreach (var itm in source)
            {
                itm?.Dispose();
            }

            source.Clear();
        }

        /// <summary>Loops through a collection of items calls IDisposable.Dispose on each item before clearing the collection</summary>
        /// <typeparam name="T">A type that implements IDisposible</typeparam>
        /// <param name="source">The collection of items to dispose and clear. Can be null.</param>
        public static void ClearAndDispose<T>(this IList<T> source) where T : IDisposable
        {
            if (source == null)
                return;

            foreach (var itm in source)
            {
                itm?.Dispose();
            }

            source.Clear();
        }


        /// <summary>
        ///     <para>Loops through a collection in reverse order performing the supplied action on each item.</para>
        ///     <para>Items in the collection can be safely removed.</para>
        /// </summary>
        /// <typeparam name="T">The Type of each item in the collection.</typeparam>
        /// <param name="source">The collection to enumerate.</param>
        /// <param name="action">The action to perform on each item.</param>
        public static void ForEachSafeReverse<T>(this IEnumerable<T> source, [NotNull] Action<T> action)
        {
            if (source == null)
                return;
            Check.NotNull(action, nameof(action));

            var listOfItems = source as IList<T> ?? new List<T>(source);

            for (var i = listOfItems.Count - 1; i >= 0; i -= 1)
            {
                action(listOfItems[i]);
            }
        }


        /// <summary>Extension method for adding a range of items to the supplied collection.</summary>
        /// <typeparam name="T">The Type of items in the collection to add.</typeparam>
        /// <param name="source">The collection to add the items to.</param>
        /// <param name="items">The items to add to the collection.</param>
        public static void AddRange<T>(this ICollection<T> source, IEnumerable<T> items)
        {
            if (source == null)
                return;
            items.ForEach(source.Add);
        }


        /// <summary>Returns all items in the collection that are deemed as duplicate by the provided selector.</summary>
        /// <typeparam name="TSource">The type of item in the source collection</typeparam>
        /// <typeparam name="TKey">The type of the key used by the selector to select the duplicates</typeparam>
        /// <param name="source">The collection to find the duplicates in</param>
        /// <param name="selector">A function used to select the key for comparing if an item is a duplicate</param>
        public static IEnumerable<TSource> Duplicates<TSource, TKey>(this IEnumerable<TSource> source,
                                                                     Func<TSource, TKey> selector)
        {
            var grouped = source.GroupBy(selector);
            var moreThanOne = grouped.Where(i => i.ContainsMultiple());
            return moreThanOne.SelectMany(i => i);
        }

        /// <summary>
        ///     Returns all items in the collection that are deemed as duplicate. The object's GetHashCode value is used as
        ///     the key to determine if an item is a duplicate
        /// </summary>
        /// <typeparam name="TSource">The type of item in the source collection</typeparam>
        /// <param name="source">The collection to find the duplicates in.</param>
        public static IEnumerable<TSource> Duplicates<TSource>(this IEnumerable<TSource> source)
        {
            return source.Duplicates(i => i);
        }


        /// <summary>Returns true if the collection contains more than one item.</summary>
        /// <typeparam name="T">The type of items in the collection.</typeparam>
        /// <param name="source">The collection of items to check.</param>
        public static bool ContainsMultiple<T>(this IEnumerable<T> source) => ContainsMinCount(source, 2, predicate: null);

        /// <summary>Returns true if the collection contains more than one item.</summary>
        /// <typeparam name="T">The type of items in the collection.</typeparam>
        /// <param name="source">The collection of items to check.</param>
        /// <param name="predicate">A predicate used to filter for specific items</param>
        public static bool ContainsMultiple<T>(this IEnumerable<T> source, Predicate<T> predicate) => ContainsMinCount(source, 2, predicate);


        /// <summary>Returns true if the collection contains more than one item.</summary>
        /// <typeparam name="T">The type of items in the collection.</typeparam>
        /// <param name="source">The collection of items to check.</param>
        public static bool ContainsSingle<T>(this IEnumerable<T> source) => ContainsMinCount(source, 1, predicate: null);

        /// <summary>Returns true if the collection contains a single item matching the passed predicate.</summary>
        /// <typeparam name="T">The type of items in the collection.</typeparam>
        /// <param name="source">The collection of items to check.</param>
        /// <param name="predicate">A predicate used to filter for specific items</param>
        public static bool ContainsSingle<T>(this IEnumerable<T> source, Predicate<T> predicate) => ContainsMinCount(source, 1, predicate);

        private static bool ContainsMinCount<T>(this IEnumerable<T> source, int minCount, Predicate<T> predicate)
        {
            if (source == null)
                return false;

            foreach (var itm in source)
            {
                if (predicate == null || predicate(itm))
                    minCount--;

                if (minCount < 0)
                    return false;
            }

            return minCount == 0;
        }


        /// <summary>Returns true if the collection contains more than one item.</summary>
        /// <typeparam name="T">The type of items in the collection.</typeparam>
        /// <param name="source">The collection of items to check.</param>
        /// <param name="item">The first and only item in the collection, or null if there are no items in the collection.</param>
        public static bool SingleOrNull<T>(this IEnumerable<T> source, out T item)
        {
            if (source == null) {
                item = default(T);
                return false;
            }


            using (var enumerator = source.GetEnumerator())
            {
                if (!enumerator.MoveNext())
                {
                    item = default(T);
                    return false;
                }

                item = enumerator.Current;
                if (!enumerator.MoveNext())
                    return true;
                 
                return false;
            }
        }


        /// <summary>A method for performing an action on each item in the collection.</summary>
        /// <typeparam name="T">The type ob object to perform the action on.</typeparam>
        /// <param name="source">The collection of items to perform the action on.</param>
        /// <param name="action">The action to perform.</param>
        public static void ForEach<T>(this IEnumerable<T> source, [NotNull] Action<T> action)
        {
            if (source == null)
                return;

            Check.NotNull(action, nameof(action));

            foreach (var itm in source)
            {
	            action(itm);
            }
        }

        /// <summary>A method for performing an action on each item in IEnumerable source.</summary>
        /// <param name="source">The collection of items to perform the action on.</param>
        /// <param name="action">The action to perform.</param>
        public static void ForEach(this IEnumerable source, Action<object> action)
        {
            if (source == null)
                return;
            Check.NotNull(action, nameof(action));
            foreach (var itm in source)
            {
                action(itm);
            }
        }


        /// <summary>
        ///     Returns the item at the top of the Stack without removing it. If the stack contains no items a null value is
        ///     returned
        /// </summary>
        /// <typeparam name="TValue">The type of items contained within the Stack. These items must a class type.</typeparam>
        /// <param name="source"></param>
        
        public static TValue TryPeek<TValue>(this Stack<TValue> source) where TValue : class
        {
            return source?.Count == 0 ? null : source.Peek();
        }
    }
}