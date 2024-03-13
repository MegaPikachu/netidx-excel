using System;
using ExcelDna.Integration;
using System.Runtime.InteropServices;
using System.Text;

namespace netidxExcelDnaAddin
{

    public class Class1 : IExcelAddIn
    {
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short write_value_f64(string path, double value, short request_type);
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short write_value_null(string path, short request_type);
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short write_value_string(string path, byte[] value, short request_type);
        
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short write_value_timestamp(string path, double value, short request_type);
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short write_value_bool(string path, bool value, short request_type);
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short write_value_i64(string path, System.Int64 value, short request_type);
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short write_value_error(string path, string value, short request_type);
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short refresh_path(string path);
        [DllImport("C:\\netidx-addin\\lib\\netidx_excel.dll")]
        static extern short refresh_all();
        static double TIMEZONE = -TimeZoneInfo.Local.GetUtcOffset(DateTime.Now).Hours/ 24.0;

        public void AutoOpen() {}

        public void AutoClose() {}

        static short try_write_auto(string path, object value, short request_type)
        {
            if (value is bool)
            {
                return write_value_bool(path, (bool)value, request_type);
            }
            else if (value is double)
            {
                return write_value_f64(path, (double)value, request_type);
            }
            else if (value is string)
            {
                return write_value_string(path, Encoding.UTF8.GetBytes((string)value), request_type);
            }
            else
            {
                return write_value_error(path, value.GetType().Name, request_type); // publish Error for unsupport values
            }
        }

        static short try_write_f64(string path, object value, short request_type)
        {
            return value as double? switch
            {
                double v => write_value_f64(path, v, request_type),
                _ => write_value_error(path, value.GetType().Name, request_type)
            };
        }

        static short try_write_i64(string path, object value, short request_type)
        {
            return value as double? switch
            {
                double v => write_value_i64(path, (System.Int64)v, request_type), // Should we write an error if [v] has a non-epsilon fractional part?
                _ => write_value_error(path, value.GetType().Name, request_type)
            };
        }

        static short try_write_null(string path, short request_type)
        {
            return write_value_null(path, request_type);
        }

        static short try_write_time(string path, object value, short request_type)
        {
            return value as double? switch
            {
                double v => write_value_timestamp(path, v + TIMEZONE, request_type),
                _ => write_value_error(path, value.GetType().Name, request_type)
            };
        }

        static short try_write_bool(string path, object value, short request_type)
        {
            return value as bool? switch
            {
                bool v => write_value_bool(path, v, request_type),
                _ => write_value_error(path, value.GetType().Name, request_type)
            };
        }

        static short try_write_string(string path, object value, short request_type)
        {
            return value as string switch
            {
                string v => write_value_string(path, Encoding.UTF8.GetBytes((string)v), request_type),
                _ => write_value_error(path, value.GetType().Name, request_type)
            };
        }
        [ExcelFunction(Description = "Write data to netidx container", IsMacroType = false, IsExceptionSafe = true, IsThreadSafe = true, IsVolatile = false)]
        public static object NetSet(string path, object value, string type = "")
        {
            short result = (short)ExcelError.ExcelErrorNA;
            short request_type = 0;
            result = WriteToNetidx(path, value, type, request_type);
            return ConvertResult(result);
        }

        [ExcelFunction(Description = "Write data to netidx container", IsMacroType = false, IsExceptionSafe = true, IsThreadSafe = true, IsVolatile = false)]
        public static object NetSetRetry(string path, object value, string type = "")
        {
            short result = (short)ExcelError.ExcelErrorNA;
            short request_type = 1;
            result = WriteToNetidx(path, value, type, request_type);
            return ConvertResult(result);
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
            return ConvertResult(result);
        }

        [ExcelFunction(Description = "Refresh subsciption of all netidx paths", IsMacroType = false, IsExceptionSafe = true, IsVolatile = false)]
        public static object RefreshAll()
        {
            short result = (short)ExcelError.ExcelErrorNA;
            result = refresh_all();
            return ConvertResult(result);
        }

        static object ConvertResult(short result)
        {
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

        static short WriteToNetidx(string path, object value, string type = "", short request_type = 0)
        {
            short result = type switch
            {
                "" => try_write_auto(path, value, request_type),
                "f64" => try_write_f64(path, value, request_type),
                "i64" => try_write_i64(path, value, request_type),
                "null" => try_write_null(path, request_type),
                "time" => try_write_time(path, value, request_type),
                "string" => try_write_string(path, value, request_type),
                "bool" => try_write_bool(path, value, request_type),
                _ => (short)ExcelError.ExcelErrorNA
            };
            return result;
        }
    }
}
