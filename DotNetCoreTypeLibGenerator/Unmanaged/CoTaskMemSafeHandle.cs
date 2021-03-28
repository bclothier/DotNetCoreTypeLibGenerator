using System;
using System.Runtime.InteropServices;

namespace DotNetCoreTypeLibGenerator.Unmanaged
{
    public class CoTaskMemSafeHandle : SafeHandle
    {
        private bool _isReleased;

        public CoTaskMemSafeHandle(int size) : base(IntPtr.Zero, true)
        {
            handle = Marshal.AllocCoTaskMem(size);
        }

        public override bool IsInvalid => handle == IntPtr.Zero;

        protected override bool ReleaseHandle()
        {
            if (handle != IntPtr.Zero && !_isReleased)
            {
                Marshal.FreeCoTaskMem(handle);
                _isReleased = true;
            }

            return true;
        }
    }
}
