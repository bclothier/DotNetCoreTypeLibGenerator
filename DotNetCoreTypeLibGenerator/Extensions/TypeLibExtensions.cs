using System;
using System.Runtime.InteropServices.ComTypes;

namespace DotNetCoreTypeLibGenerator.Extensions
{
    public static class TypeLibExtensions
    {
        public static void UsingLibAttr(this ITypeLib typeLib, Action<TYPELIBATTR> action)
        {
            ExtensionHelper.UsingPtrToStructure(ptr => { typeLib.GetLibAttr(out ptr); return ptr; }, action, typeLib.ReleaseTLibAttr);
        }
    }
}
