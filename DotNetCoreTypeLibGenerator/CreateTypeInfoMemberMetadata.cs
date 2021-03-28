using DotNetCoreTypeLibGenerator.Helpers;
using System.Reflection;
using System.Runtime.InteropServices;

namespace DotNetCoreTypeLibGenerator
{
    internal readonly struct CreateTypeInfoMemberMetadata
    {
        internal VarEnum VarEnum { get; }
        internal int MemberId { get; }
        internal int? DispId { get; }
        internal MemberInfo MemberInfo { get; }

        public CreateTypeInfoMemberMetadata(int memberId, MemberInfo memberInfo, VarEnum varEnum)
        {
            MemberId = memberId;
            MemberInfo = memberInfo;
            DispId = AttributeHelper.GetDispId(MemberInfo);
            VarEnum = varEnum;
        }
    }
}
