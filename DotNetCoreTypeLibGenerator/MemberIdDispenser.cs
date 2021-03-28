using System.Collections.Generic;

namespace DotNetCoreTypeLibGenerator
{
    /// <summary>
    /// Encapsulates the assignment of MEMBERID to ensure that a member within a single type will always have the
    /// same MEMBERID. A type may have property accessors that requires re-use of the same MEMBERID, so we must
    /// ensure that when same member name is provided, the internal counter do not increment. 
    /// 
    /// See MS-OAUT Section 2.2.35 - MEMBERID
    /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/ace8758f-ee2b-4cb6-8645-973994d12530
    /// </summary>
    internal class MemberIdDispenser
    {
        private readonly Dictionary<string, int> _memberIds = new Dictionary<string, int>();
        private int _memberId = 0;
        
        public int GetMemberId(string memberName)
        {
            if (_memberIds.TryGetValue(memberName, out var existingMemberId))
            {
                return existingMemberId;
            }

            var assignedMemberId = _memberId++;
            _memberIds.Add(memberName, assignedMemberId);
            return assignedMemberId;
        }
    }
}
