using DotNetCoreTypeLibGenerator.Unmanaged;
using System;
using System.Runtime.InteropServices;
using static Vanara.PInvoke.OleAut32;

using PARAMDESC = System.Runtime.InteropServices.ComTypes.PARAMDESC;

namespace DotNetCoreTypeLibGenerator
{
    /// <summary>
    /// Described in MS-OAUT section 2.2.39. 
    /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/683c767d-2e8e-4d2f-8804-afeb3a73969a
    /// Not implemented in either the <see cref="System.Runtime.InteropServices.ComTypes"/>
    /// nor in <see cref="Vanara.PInvoke.OleAut32"/>. Because the <see cref="PARAMDESC.lpVarValue"/> takes a 
    /// <see cref="IntPtr"/>, we must allow for converting the <see cref="PARAMDESCEX"/> into a pointer.
    /// </summary>
    [
        StructLayout(LayoutKind.Sequential)
    ]
    internal readonly struct PARAMDESCEX
    {
        /// <summary>
        /// Should be the size of the <see cref="PARAMDESCEX"/> structure. See footnote 20 in the section 2.2.39
        /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/ae19ad63-7433-4568-88e9-f70e5593547d#Appendix_A_20
        /// </summary>
        public uint cBytes { get => (uint)Marshal.SizeOf<PARAMDESCEX>(); }
        public VARIANT varDefaultValue { get; }

        public PARAMDESCEX(object value)
        {
            VARIANT tmp = new VARIANT();
            UnmanagedMemory.UsingVariant(value, ptr => {
                tmp = (VARIANT)Marshal.PtrToStructure(ptr, typeof(VARIANT));
            });
            varDefaultValue = tmp;
        }

        /// <returns>
        /// Returns a unmanaged pointer representing the <see cref="PARAMDESCEX"/> structure. The caller is now 
        /// responsible for freeing the pointer and should call <see cref="SafeHandle.ReleaseHandle"/> on it
        /// when done with it. 
        /// </returns>
        public SafeHandle GetHandle()
        {
            var handle = new CoTaskMemSafeHandle(Marshal.SizeOf(this));
            Marshal.StructureToPtr(this, handle.DangerousGetHandle(), false);
            return handle;
        }
    }
}
