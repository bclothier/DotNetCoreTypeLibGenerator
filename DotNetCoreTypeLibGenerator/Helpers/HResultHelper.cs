using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Vanara.PInvoke;
using static Vanara.PInvoke.OleAut32;

namespace DotNetCoreTypeLibGenerator.Helpers
{
    public static class HResultHelper
    {
        public static void SUCCEEDED(HRESULT hr)
        {
            if(hr.Failed)
            {
                throw new COMException(hr.FormatMessage(), hr.Code);
            }
        }
    }
}
