using System;
using System.Runtime.InteropServices;
using static Vanara.PInvoke.OleAut32;

namespace DotNetCoreTypeLibGenerator.Unmanaged
{
    /// <summary>
    /// Helps encapsulates the unmanaged access by simulating a using block and is useful
    /// for unmanaged memory allocation/deallocation within a local scope. For allocation
    /// that needs to persist outside a local scope, use <see cref="CoTaskMemSafeHandle"/> instead.
    /// </summary>
    public static class UnmanagedMemory
    {
        public static void UsingCoTaskMem(int size, Action<IntPtr> action)
        {
            IntPtr ptr = IntPtr.Zero;
            try
            {
                if (size != 0)
                {
                    ptr = Marshal.AllocCoTaskMem(size);
                }
                action.Invoke(ptr);
            }
            finally
            {
                if (ptr != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(ptr);
                }
            }
        }

        public static void UsingVariant(object value, Action<IntPtr> action)
        {
            UsingCoTaskMem(Marshal.SizeOf<VARIANT>(), ptr =>
            {
                Marshal.GetNativeVariantForObject(value, ptr);
                action.Invoke(ptr);
            });
        }
    }
}
