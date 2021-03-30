using DotNetCoreTypeLibGenerator.Unmanaged;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace DotNetCoreTypeLibGenerator.Extensions
{
    public static class TypeDescExtensions
    {
        /// <returns>
        /// Returns a unmanaged pointer representing the <see cref="TYPEDESC"/> structure. The caller is now 
        /// responsible for freeing the pointer and should call <see cref="SafeHandle.ReleaseHandle"/> on it
        /// when done with it. 
        /// </returns>
        public static SafeHandle GetHandle(this TYPEDESC typeDesc)
        {
            var handle = new CoTaskMemSafeHandle(Marshal.SizeOf(typeDesc));
            Marshal.StructureToPtr(typeDesc, handle.DangerousGetHandle(), false);
            return handle;
        }
    }
}
