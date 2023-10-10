using System;
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using System.Text;

namespace netidxExcelDnaAddin
{
    public class Class1 : IExcelAddIn
    {
        [DllImport("netidx_excel.dll")]
        static extern short write_value_f64(string path, double value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_string(string path, byte[] value);
        
        [DllImport("netidx_excel.dll")]
        static extern short write_value_timestamp(string path, double value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_bool(string path, bool value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_i64(string path, System.Int64 value);
        [DllImport("netidx_excel.dll")]
        static extern short write_value_error(string path, string value);
        [DllImport("netidx_excel.dll")]
        static extern short refresh_path(string path);
        [DllImport("netidx_excel.dll")]
        static extern short refresh_all();
        static double TIMEZONE = -TimeZoneInfo.Local.GetUtcOffset(DateTime.Now).Hours/ 24.0;

        public void AutoOpen() { }

        public void AutoClose() { }

        static short try_write_auto(string path, object value)
        {
            if (value is bool)
            {
                return write_value_bool(path, (bool)value);
            }
            else if (value is double)
            {
                return write_value_f64(path, (double)value);
            }
            else if (value is string)
            {
                return write_value_string(path, Encoding.UTF8.GetBytes((string)value));
            }
            else
            {
                return write_value_error(path, value.GetType().Name); // publish Error for unsupport values
            }
        }

        static short try_write_f64(string path, object value)
        {
            return value as double? switch
            {
                double v => write_value_f64(path, v),
                _ => write_value_error(path, value.GetType().Name)
            };
        }

        static short try_write_i64(string path, object value)
        {
            return value as double? switch
            {
                double v => write_value_i64(path, (System.Int64)v), // Should we write an error if [v] has a non-epsilon fractional part?
                _ => write_value_error(path, value.GetType().Name)
            };
        }

        static short try_write_time(string path, object value)
        {
            return value as double? switch
            {
                double v => write_value_timestamp(path, v + TIMEZONE),
                _ => write_value_error(path, value.GetType().Name)
            };
        }

        static short try_write_bool(string path, object value)
        {
            return value as bool? switch
            {
                bool v => write_value_bool(path, v),
                _ => write_value_error(path, value.GetType().Name)
            };
        }

        static short try_write_string(string path, object value)
        {
            return value as string switch
            {
                string v => write_value_string(path, Encoding.UTF8.GetBytes((string)v)),
                _ => write_value_error(path, value.GetType().Name)
            };
        }

        [ExcelFunction(Description = "Write data to netidx container", IsMacroType = false, IsExceptionSafe = true, IsThreadSafe = true, IsVolatile = false)]
        public static object NetSet(string path, object value, string type = "")
        {
            short result = (short)ExcelError.ExcelErrorNA;
            result = type switch
            {
                "" => try_write_auto(path, value),
                "f64" => try_write_f64(path, value),
                "i64" => try_write_i64(path, value),
                "time" => try_write_time(path, value),
                "string" => try_write_string(path, value),
                "bool" => try_write_bool(path, value),
                _ => (short)ExcelError.ExcelErrorNA
            };
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

        [ExcelFunction(Description = "Refresh subsciption of a netidx path", IsMacroType = false, IsExceptionSafe = true, IsVolatile = false)]
        public static object RefreshPath(string path)
        {
            short result = (short)ExcelError.ExcelErrorNA;
            result = refresh_path(path);
            switch (result)
            {
                case -1:
                    return "#SET";
                case -2:
                    return "#MAYBE_SET";
                default:
                    return (ExcelError)result;
            }
        }

        [ExcelFunction(Description = "Refresh subsciption of all netidx paths", IsMacroType = false, IsExceptionSafe = true, IsVolatile = false)]
        public static object RefreshAll()
        {
            short result = (short)ExcelError.ExcelErrorNA;
            result = refresh_all();
            switch (result)
            {
                case -1:
                    return "#SET";
                case -2:
                    return "#MAYBE_SET";
                default:
                    return (ExcelError)result;
            }
        }
    }
}
