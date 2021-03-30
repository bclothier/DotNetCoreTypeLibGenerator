using System.Runtime.InteropServices;

namespace Structs
{
    [
        ComVisible(true),
        Guid("2F67AD81-2EF2-4E9A-8122-6FB1DFAB9CA2")
    ]
    public struct SimpleStruct
    {
        [DispId(23)]
        public int SomeInteger;
        [DispId(42)]
        public double SomeDouble;
    }
}
