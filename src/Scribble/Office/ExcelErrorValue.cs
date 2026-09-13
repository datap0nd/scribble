using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Scribble.Office
{
    // Value2 can marshal VT_ERROR as a CVErr, signed HRESULT, or ErrorWrapper.
    public static class ExcelErrorValue
    {
        private static readonly Dictionary<int, string> Names = new Dictionary<int, string> {
            { 2000, "#NULL!" }, { 2007, "#DIV/0!" }, { 2015, "#VALUE!" },
            { 2023, "#REF!" }, { 2029, "#NAME?" }, { 2036, "#NUM!" },
            { 2042, "#N/A" }, { 2045, "#SPILL!" }
        };

        public static string Text(object value)
        {
            var wrapper = value as ErrorWrapper;
            if (wrapper != null) value = wrapper.ErrorCode;
            if (!(value is int)) return null;
            var code = (int)value;
            var hresult = (unchecked((uint)code) & 0xffff0000u) == 0x800a0000u;
            if (hresult) code &= 0xffff;
            string text;
            return Names.TryGetValue(code, out text) ? text :
                (wrapper != null || (hresult && code >= 2000 && code < 2100)) ? "#EXCEL_ERROR(" + code + ")" : null;
        }
    }
}
