// © Copyright 2018 Levit & James, Inc.

using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace LevitJames.Interop
{
    internal class SafeHGlobalHandle : SafeHandle
    {
        [SuppressMessage("Microsoft.Reliability", "CA2006:UseSafeHandleToEncapsulateNativeResources")]
        private IntPtr
            _data;
#if (TRACK_DISPOSED)
        private readonly string _disposedSource;
#endif

        public SafeHGlobalHandle(IntPtr handle, bool ownsHandle) : base(IntPtr.Zero, ownsHandle)
        {
            SetHandle(handle);
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        public SafeHGlobalHandle(IntPtr handle) : base(IntPtr.Zero, ownsHandle: true)
        {
            SetHandle(handle);
#if (TRACK_DISPOSED)
            _disposedSource = LevitJames.Core.Diagnostics.DisposeTracker.GetSource();
#endif
        }

        public override bool IsInvalid => DangerousGetHandle() == IntPtr.Zero;

        public IntPtr LockedData => _data;

        [SuppressMessage("Microsoft.Naming", "CA2204:Literals should be spelled correctly",
            MessageId = "SafeHGlobalHandle")]
        public int GetSize()
        {
            if (!IsInvalid)
            {
                if (!IsLocked)
                {
                    throw new InvalidOperationException("SafeHGlobalHandle not locked");
                }

                return NativeMethods.GlobalSize(LockedData);
            }

            return 0;
        }

        public bool Lock()
        {
            if (IsInvalid)
            {
                return false;
            }

            if (_data == IntPtr.Zero)
            {
                _data = NativeMethods.GlobalLock(this);
            }

            return IsLocked;
        }

        public bool IsLocked => _data != IntPtr.Zero;

        public bool Unlock()
        {
            if (!IsInvalid && _data == IntPtr.Zero)
            {
                _data = IntPtr.Zero;
                return NativeMethods.GlobalUnlock(this);
            }

            return false;
        }

        protected override bool ReleaseHandle()
        {
            if (!IsInvalid)
            {
                if (IsLocked)
                    Unlock();

                var ret = NativeMethods.GlobalFree(DangerousGetHandle());
                var lastError = Marshal.GetLastWin32Error();
                if (ret == IntPtr.Zero)
                {
                    SetHandleAsInvalid();
                    return true;
                }

                SetHandleAsInvalid();
                throw new Win32Exception(lastError);
            }

            return false;
        }

        ~SafeHGlobalHandle()
        {
#if (TRACK_DISPOSED)
                LevitJames.Core.Diagnostics.DisposeTracker.Add(_disposedSource);
#endif
            Dispose(false);
        }
    }
}