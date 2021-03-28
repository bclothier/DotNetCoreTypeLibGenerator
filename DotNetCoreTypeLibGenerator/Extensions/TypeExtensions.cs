using System;
using System.Runtime.InteropServices;
using VARENUM = System.Runtime.InteropServices.VarEnum;

namespace DotNetCoreTypeLibGenerator.Extensions
{
    public static class TypeExtensions
    {
        /// <summary>
        /// Helps infers the <see cref="VARENUM"/> for a given .NET type.
        /// Must conform to the MS-OAUT section 2.2.7 VARIANT Type constants. See:
        /// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/3fe7db9f-5803-4dc4-9d14-5425d3f5461f
        /// </summary>
        /// <param name="type">A .NET type to be marshaled to COM.</param>
        /// <returns>A suitable <see cref="VARENUM"/> type for the variants containing this .NET type.</returns>
        public static VARENUM GetVarEnum(this Type type)
        {
            // TODO: Check whether boxed variables need to be handled.
            // TODO: Fix the logic for the VT_BYREF flag for reference types, requires consumers to also handle it accordingly
            // TODO: Fix the logic for the VT_ARRAY flag if the type is an array type, requires consumers to also handle it accordingly
            // TODO: Write tests for this method for all path

            VARENUM vt = VARENUM.VT_EMPTY;

            //VT_VOID must be checked first to avoid mismapping it. 
            if (type == typeof(void)) return VARENUM.VT_VOID;

            //VT_ARRAY
            //if (type.IsArray) vt |= VARENUM.VT_ARRAY;
            //VT_BYREF
            //if (!type.IsValueType) vt |= VARENUM.VT_BYREF;

            //VT_BOOL
            if (type == typeof(bool)) return vt |= VARENUM.VT_BOOL;
            //VT_BSTR
            if (type == typeof(string)) return vt |= VARENUM.VT_BSTR;
            //VT_CY
            if (type == typeof(CurrencyWrapper)) return vt |= VARENUM.VT_CY;
            //VT_DATE
            if (type == typeof(DateTime)) return vt |= VARENUM.VT_DATE;
            //VT_DECIMAL
            if (type == typeof(decimal)) return vt |= VARENUM.VT_DECIMAL;
            //VT_ERROR
            if (type == typeof(ErrorWrapper)) return vt |= VARENUM.VT_ERROR;
            //VT_I1
            if (type == typeof(sbyte)) return vt |= VARENUM.VT_I1;
            //VT_I2
            if (type == typeof(short)) return vt |= VARENUM.VT_I2;
            //VT_I4, VT_INT
            if (type == typeof(int)) return vt |= VARENUM.VT_I4;
            //VT_I8
            if (type == typeof(long)) return vt |= VARENUM.VT_I8;
            //VT_NULL
            if (type == typeof(DBNull)) return vt |= VARENUM.VT_NULL;
            //VT_R4
            if (type == typeof(float)) return vt |= VARENUM.VT_R4;
            //VT_R8
            if (type == typeof(double)) return vt |= VARENUM.VT_R8;
            //VT_RECORD
            if (type.IsValueType && !type.IsPrimitive) return vt |= VARENUM.VT_RECORD;
            //VT_UI1
            if (type == typeof(byte)) return vt |= VARENUM.VT_UI1;
            //VT_UI2
            if (type == typeof(ushort)) return vt |= VARENUM.VT_UI2;
            //VT_UI4, VT_UINT
            if (type == typeof(uint)) return vt |= VARENUM.VT_UI4;
            //VT_UI8
            if (type == typeof(ulong)) return vt |= VARENUM.VT_UI8;
            //VT_UNKNOWN
            if (type.IsClass || type.IsInterface) return vt |= VARENUM.VT_UNKNOWN;
            
            //VT_VARIANT
            if (type.IsValueType && type.IsPrimitive) return vt |= VARENUM.VT_VARIANT;
            //VT_DISPATCH
            if (type == typeof(object)) return vt |= VARENUM.VT_DISPATCH;

            throw new InvalidCastException($"Cannot cast the type {type.Name} to a valid VARENUM.");
        }
    }
}
