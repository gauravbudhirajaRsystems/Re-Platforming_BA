// © Copyright 2018 Levit & James, Inc.

using System;
using System.Diagnostics;
namespace LevitJames.Core.Common
{
    /// <summary>
    ///     A simple structure for managing and maintaining a lock counter.
    /// </summary>
    /// <remarks>
    ///     A lock counter is structure that keeps track of the number of Lock and Unlock calls. The LockCounter is only
    ///     in an unlocked state the counter is 0. i.e.
    ///     when there have been the same number of Lock and UnLock calls.
    /// </remarks>
    [Serializable]
    [DebuggerStepThrough]
    public class LockCounter
    {
        /// <summary>
        ///     Returns/sets the locked state of the object
        /// </summary>

        /// <returns>True if the object is locked false otherwise.</returns>

        public bool Locked => LockCount > 0;

        /// <summary>
        ///     Returns the number of Locks on this object
        /// </summary>



        public int LockCount { get; private set; }

        /// <summary>
        ///     Increments the LockCounter
        /// </summary>
        /// <returns>True if this is the first call to Lock, i.e. the on return the LockCount = 1;otherwise false</returns>
        /// <remarks>To unlock the object there must be the same number of calls to Unlock or a single call to ClearAllLocks</remarks>
        public bool Lock()
        {
            LockCount += 1;
            return LockCount == 1;
        }

        /// <summary>
        ///     Decrements the LockCounter
        /// </summary>
        /// <returns>Returns true if the call unlocked the counter, i.e. on return the LockCount has been set to 0 ;otherwise false</returns>
        /// <remarks>If you unlock more times that you Lock then the return value is always 0</remarks>
        public bool Unlock()
        {
            if (LockCount == 0)
            {
                return false;
            }

            LockCount -= 1;
            return !Locked;
        }

        /// <summary>
        ///     Resets the LockCounter back to zero
        /// </summary>
        public void Reset()
        {
            LockCount = 0;
        }
    }
}
