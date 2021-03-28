using System;
using System.Runtime.InteropServices.ComTypes;

namespace DotNetCoreTypeLibGenerator.Extensions
{
    public static class TypeInfoExtensions
    {
        public static void UsingTypeAttr(this ITypeInfo typeInfo, Action<TYPEATTR> action)
        {
            ExtensionHelper.UsingPtrToStructure(ptr => { typeInfo.GetTypeAttr(out ptr); return ptr; }, action, typeInfo.ReleaseTypeAttr);
        }

        public static void UsingVarDesc(this ITypeInfo typeInfo, int memId, Action<VARDESC> action)
        {
            ExtensionHelper.UsingPtrToStructure(ptr => { typeInfo.GetVarDesc(memId, out ptr); return ptr; }, action, typeInfo.ReleaseVarDesc);
        }

        public static void UsingFuncDesc(this ITypeInfo typeInfo, int memId, Action<FUNCDESC> action)
        {
            ExtensionHelper.UsingPtrToStructure(ptr => { typeInfo.GetFuncDesc(memId, out ptr); return ptr; }, action, typeInfo.ReleaseFuncDesc);
        }
    }
}
