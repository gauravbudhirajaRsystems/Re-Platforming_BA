using System;
using System.Runtime.InteropServices;
 
namespace LevitJames.Win32
{
    [StructLayout(LayoutKind.Sequential)]
    public struct DpiAwarenessContextHandle
    {
 
        private readonly IntPtr _handle;
 
        public IntPtr Handle => _handle;

        public DpiAwareness Resolve()
        {
            var context = new DpiAwarenessContextHandle(DpiAwareness.SystemAware);
            if (NativeMethods.AreDpiAwarenessContextsEqual(this, context))
                return DpiAwareness.SystemAware;

            context = new DpiAwarenessContextHandle(DpiAwareness.PerMonitorAware2);
            if (NativeMethods.AreDpiAwarenessContextsEqual(this, context))
                return DpiAwareness.PerMonitorAware2;
 
            context = new DpiAwarenessContextHandle(DpiAwareness.PerMonitorAware);
            if (NativeMethods.AreDpiAwarenessContextsEqual(this, context))
                return DpiAwareness.PerMonitorAware;

            context = new DpiAwarenessContextHandle(DpiAwareness.Unaware);
            if (NativeMethods.AreDpiAwarenessContextsEqual(this, context))
                return DpiAwareness.Unaware;

            return DpiAwareness.Unaware;
        }

        private DpiAwarenessContextHandle(IntPtr handle)
        {
            _handle = handle;
        }

        public DpiAwarenessContextHandle(DpiAwareness awareness)
        {
            switch (awareness)
            {
            case DpiAwareness.Unaware:
                _handle = (IntPtr)NativeMethods.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_UNAWARE;
                break;
            case DpiAwareness.SystemAware:
                _handle = (IntPtr)NativeMethods.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_SYSTEM_AWARE;
                break;
            case DpiAwareness.PerMonitorAware:
                _handle = (IntPtr)NativeMethods.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE;
                break;
            case DpiAwareness.PerMonitorAware2:
                _handle = (IntPtr)NativeMethods.DPI_AWARENESS_CONTEXT.DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(awareness), awareness, null);
            }
        }

        /// <summary>Returns the fully qualified type name of this instance.</summary>
        /// <returns>The fully qualified type name.</returns>
        public override string ToString()
        {
            return $"Handle:={Handle}, ({Resolve()})" ;
        }
        
        public static DpiAwarenessContextHandle FromThread => NativeMethods.GetThreadDpiAwarenessContext();
        public static DpiAwarenessContextHandle FromWindow(IntPtr handle) => NativeMethods.GetWindowDpiAwarenessContext(new HandleRef(null, handle));

        public static bool IsProcessDPIAware() => NativeMethods.IsProcessDPIAware();
 
        public static implicit operator IntPtr(DpiAwarenessContextHandle source) => source.Handle;
        public static implicit operator DpiAwarenessContextHandle(IntPtr source) => new DpiAwarenessContextHandle(source);

        public static implicit operator DpiAwarenessContextHandle(DpiAwareness source) => new DpiAwarenessContextHandle(source);
 
        public static implicit operator DpiAwareness(DpiAwarenessContextHandle source) => source.Resolve();

        /// <summary>Indicates whether this instance and a specified object are equal.</summary>
        /// <param name="obj">The object to compare with the current instance. </param>
        /// <returns>
        /// <see langword="true" /> if <paramref name="obj" /> and this instance are the same type and represent the same value; otherwise, <see langword="false" />. </returns>
        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (obj is DpiAwareness contextAsEnum)
                return Equals(this, new DpiAwarenessContextHandle(contextAsEnum));

            if (obj is int contextAsInt)
                return Equals(this, new DpiAwarenessContextHandle((DpiAwareness)contextAsInt));

            if (!(obj is DpiAwarenessContextHandle context))
                return false;

            return Equals(context);
        }

        public bool Equals(DpiAwarenessContextHandle obj) => NativeMethods.AreDpiAwarenessContextsEqual(this, obj);

        /// <inheritdoc />
        public override int GetHashCode()
        {
            return _handle.ToInt32();
        }
    }
}