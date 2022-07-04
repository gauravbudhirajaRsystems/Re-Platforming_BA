// © Copyright 2018 Levit & James, Inc.

using System;
using System.Collections.Generic;

namespace LevitJames.Core
{
    /// <summary>
    ///     A class for storing the set of differences between two serialized streams
    /// </summary>
    public struct SerializationStreamComparerSet
    {
        /// <summary>
        ///     Instance1
        /// </summary>
        public object Instance1 { get; internal set; }

        /// <summary>
        ///     Instance2
        /// </summary>
        public object Instance2 { get; internal set; }

        /// <summary>
        ///     The serialized differences between Instance1 and instance2.
        /// </summary>
        public IEnumerable<SerizaliationStreamComparerItem> Differences { get; internal set; }

        /// <summary>
        ///     Calls Dispose on Instance1 and Instance2, if they implememt IDisposable.
        /// </summary>
        public void DisposeInstances()
        {
            (Instance1 as IDisposable)?.Dispose();
            (Instance2 as IDisposable)?.Dispose();
        }
    }
}