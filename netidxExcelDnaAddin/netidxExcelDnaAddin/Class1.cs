using System;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace netidxExcelDnaAddin
{
    public class Class1 : IExcelAddIn
    {
        [DllImport("netidx_excel.dll")]
        static extern short write_value_string(string path, string value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_int(string path, int value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_float(string path, double value);

        public void AutoOpen() { }

        public void AutoClose() { }

        [ExcelFunction(Description = "Write data to netidx container", IsMacroType = false, IsExceptionSafe = true, IsThreadSafe = true, IsVolatile = false)]
        public static object NetSet(string path, object value)
        {
            short result = (short)ExcelError.ExcelErrorNA;
            if (value is int)
            {
                result = write_value_int(path, (int)value);
            }
            else if (value is double || value is float)
            {
                result = write_value_float(path, (double)value);
            }
            else if (value is string)
            {
                result = write_value_string(path, (string)value);
            }
            switch(result) {
                case -1:
                    return "#SET";
                case -2:
                    return "#MAYBE_SET";
                default:
                    return (ExcelError)result;
            }
        }

        [ExcelFunction(Description = "Subscribe to a netidx path", IsMacroType = false, IsExceptionSafe = true, IsThreadSafe = true, IsVolatile = false)]
        public static object NetGet(string path)
        {
            return XlCall.RTD("netidxrtd", null, path);
        }
    }
}
