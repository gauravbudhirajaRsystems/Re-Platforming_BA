// © Copyright 2018 Levit & James, Inc.

using System;

namespace LevitJames.AddinApplicationFramework
{
    /// <summary>
    ///     A class for generating a unique id used for marking an object as dirty.
    /// </summary>
    public static class MarkAsDirtyGenerator
    {
        /// <summary>
        ///     Marks the passed IDirty instance as dirty provided the IDirty.DirtyCookie member does not equal -1.
        ///     If IDirty.DirtyCookie equals -1 then the dirty state of the object is determined by comparing the previous stream.
        /// </summary>
        /// <param name="data"></param>
        public static void MarkAsDirty(IAddinAppDirty data)
        {
            if (data == null || data.DirtyCookie != 0)
                return;

            data.DirtyCookie = Guid.NewGuid().GetHashCode();
        }
    }
}