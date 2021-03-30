using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace DotNetCoreTypeLibGenerator.Helpers
{
    /// <summary>
    /// Simplifies the access to attributes that may be decorated on a .NET type. Typically we want the
    /// value of the attribute, rather than the attribute itself and we need to also handle the case where
    /// the attribute may not exist on the type. 
    /// </summary>
    internal static class AttributeHelper
    {
        public static ComInterfaceType? GetComInterface(Type type)
        {
            return GetFirstOrDefaultAttribute<InterfaceTypeAttribute>(type)?.Value;
        }

        public static Guid GetGuid(ICustomAttributeProvider source)
        {
            return Guid.Parse(GetFirstOrDefaultAttribute<GuidAttribute>(source)?.Value);
        }

        public static bool GetComVisible(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<ComVisibleAttribute>(source)?.Value ?? false;
        }

        public static string GetComAliasName(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<ComAliasNameAttribute>(source)?.Value;
        }

        public static int? GetDispId(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<DispIdAttribute>(source)?.Value;
        }

        public static int? GetDispId(Type enumType, string enumMemberName)
        {
            return GetEnumValueAttribute<DispIdAttribute>(enumType, enumMemberName)?.Value;
        }

        public static bool IsInParameter(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<InAttribute>(source) != null;
        }

        public static bool IsOutParameter(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<OutAttribute>(source) != null;
        }

        public static bool IsOptionalParameter(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<OptionalAttribute>(source) != null;
        }

        public static object HasDefaultValue(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<DefaultParameterValueAttribute>(source)?.Value;
        }

        public static bool HasPreserveSig(ICustomAttributeProvider source)
        {
            return GetFirstOrDefaultAttribute<PreserveSigAttribute>(source) != null;
        }

        private static T GetFirstOrDefaultAttribute<T>(ICustomAttributeProvider source) where T : Attribute
        {
            var attributes = source.GetCustomAttributes(typeof(T), false);
            return attributes.Length == 1 ? (T)source.GetCustomAttributes(typeof(T), false)[0] : null;
        }

        public static T GetEnumValueAttribute<T>(Enum enumVal) where T : Attribute
        {
            var enumType = enumVal.GetType();
            var enumMemberName = enumVal.ToString();
            return GetEnumValueAttribute<T>(enumType, enumMemberName);
        }

        public static T GetEnumValueAttribute<T>(Type enumType, string enumMemberName) where T : Attribute
        {
            var memInfo = enumType.GetMember(enumMemberName);
            var attributes = memInfo[0].GetCustomAttributes(typeof(T), false);
            return attributes.Length > 0 ? (T)attributes[0] : null;
        }
    }
}
