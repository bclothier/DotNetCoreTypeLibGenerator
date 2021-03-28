using System;
using System.Runtime.InteropServices;

namespace EmptyTypeLib
{
    [
        ComVisible(true),
        Guid("DDC35805-5F49-441D-BD2F-856896564CA8")
    ]
    public enum EmptyEnum
    {
        [DispId(1)] Apple = 3,
        [DispId(3)] Orange = 2,
        [DispId(2)] Tomato = 1
    }
}
