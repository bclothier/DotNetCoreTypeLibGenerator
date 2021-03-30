using System;
using System.Runtime.InteropServices;

namespace Constants
{
    [
        ComVisible(true),
        Guid("4022645C-9D37-48A5-A80E-2AF952CF5506")
    ]
    public static class Constants
    {
        //VT_ARRAY
        //VT_BOOL
        public const bool I_AM_VT_BOOL = true;
        //VT_BSTR
        public const string I_AM_VT_BSTR = "Yes, you are!";
        //VT_BYREF

        //VT_CY
        public static readonly CurrencyWrapper I_AM_VT_CY = new CurrencyWrapper(42.56M);
        //VT_DATE
        public static readonly DateTime I_AM_VT_DATE = new DateTime(1970, 1, 1);
        //VT_DECIMAL
        public const decimal I_AM_VT_DECIMAL = 1.2M;

        //VT_DISPATCH
        //VT_EMPTY
        //VT_ERROR
        public static readonly ErrorWrapper I_AM_VT_ERROR = new ErrorWrapper(1);
        //VT_I1
        public const sbyte I_AM_VT_I1 = 1;
        //VT_I2
        public const short I_AM_VT_I2 = 2;
        //VT_I4
        public const int I_AM_VT_I4 = 4;
        //VT_INT
        //public const int I_AM_VT_INT = 5;
        //VT_I8
        public const long I_AM_VT_I8 = 8;
        //VT_NULL
        //public static readonly DBNull I_AM_VT_NULL = DBNull.Value;
        //VT_R4
        public const float I_AM_VT_R4 = 2.3F;
        //VT_R8
        public const double I_AM_VT_R8 = 3.4;
        //VT_RECORD
        //VT_UI1
        public const byte I_AM_VT_UI1 = 1;
        //VT_UI2
        public const ushort I_AM_VT_UI2 = 2;
        //VT_UI4
        public const uint I_AM_VT_UI4 = 4;
        //VT_UINT
        //public const uint I_AM_VT_UINT = 4;
        //VT_UI8
        public const ulong I_AM_VT_UI8 = 16;
        //VT_UNKNOWN
        public static readonly object I_AM_VT_UNKNOWN = 42;
        //VT_VARIANT
    }
}
