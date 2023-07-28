using System;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace netidxExcelDnaAddin
{
    public class Class1 : IExcelAddIn
    {
        [DllImport("netidx_excel.dll")]
        static extern short write_value_float(string path, double value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_string(string path, string value);
        
        [DllImport("netidx_excel.dll")]
        static extern short write_value_timestamp(string path, double value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_bool(string path, double value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_int(string path, int value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_error(string path, string value);
        static double TIMEZONE = -TimeZoneInfo.Local.GetUtcOffset(DateTime.Now).Hours/ 24.0;

        public void AutoOpen() { }

        public void AutoClose() { }

        [ExcelFunction(Description = "Write data to netidx container", IsMacroType = false, IsExceptionSafe = true, IsThreadSafe = true, IsVolatile = false)]
        public static object NetSet(string path, object value, string type)
        {
            short result = (short)ExcelError.ExcelErrorNA;
            if (type == "double" && value is double)
            {
                result = write_value_float(path, (double)value);
            }
            else if (type == "string" && value is string)
            {
                result = write_value_string(path, (string)value);
            }
            else if (type == "timestamp" && value is double)
            {
                result = write_value_timestamp(path, (double)value + TIMEZONE);
            }
            else if (type == "bool" && value is double)
            {
                result = write_value_bool(path, (double)value);
            }
            else if (type == "int" && value is double)  // we do not get int from Excel, only get double
            {
                result = write_value_int(path, (int)(double)value);
            }
            else
            {
                result = write_value_error(path, value.GetType().Name); // publish Error for unsupport values
                return (ExcelError)result;
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
