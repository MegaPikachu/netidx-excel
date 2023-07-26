using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Reflection;
using System.IO;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Resources;

namespace netidxExcelDnaAddin
{
    public class Class1 : IExcelAddIn
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr LoadLibrary(string path);

        [DllImport("kernel32.dll")]
        static extern void FreeLibrary(IntPtr lib);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetProcAddress(IntPtr lib, string proc);

        IntPtr netidx_excel_dll;

        private delegate short WriteValueString(string path, string value);
        private delegate short WriteValueInt(string path, int value);
        private delegate short WriteValueFloat(string path, double value);
        static WriteValueString write_value_string;
        static WriteValueInt write_value_int;
        static WriteValueFloat write_value_float;

        public void AutoOpen()
        {
            byte[] ba;
            Assembly curAsm = Assembly.GetExecutingAssembly();
            string resource_name = "netidx_excel.dll";

            using (Stream stm = curAsm.GetManifestResourceStream(resource_name))
            {
                ba = new byte[(int)stm.Length];
                stm.Read(ba, 0, (int)stm.Length);
                    
            }
                
            string dll_dir;
            string dll_path;
            bool write_to_disk = false;
                
            using (SHA1CryptoServiceProvider sha1 = new SHA1CryptoServiceProvider())
            {
                string fileHash = BitConverter.ToString(sha1.ComputeHash(ba));

                dll_dir = Path.GetTempPath() + "netidx_excel." + fileHash + ".dll";
                Directory.CreateDirectory(dll_dir);
                dll_path = dll_dir + "/netidx_excel.dll";

                if (File.Exists(dll_path))
                {
                    byte[] bb = File.ReadAllBytes(dll_path);
                    string fileHash2 = BitConverter.ToString(sha1.ComputeHash(bb));

                    if (fileHash != fileHash2)
                    {
                        write_to_disk = true;
                    }
                }
                else
                {
                    write_to_disk = true;
                }
            }

            if (write_to_disk)
            {
                File.WriteAllBytes(dll_path, ba);
            }

            netidx_excel_dll = LoadLibrary(dll_path);
            write_value_string = (WriteValueString)Marshal.GetDelegateForFunctionPointer(GetProcAddress(netidx_excel_dll, "write_value_string"), typeof(WriteValueString));
            write_value_int = (WriteValueInt)Marshal.GetDelegateForFunctionPointer(GetProcAddress(netidx_excel_dll, "write_value_int"), typeof(WriteValueInt));
            write_value_float = (WriteValueFloat)Marshal.GetDelegateForFunctionPointer(GetProcAddress(netidx_excel_dll, "write_value_float"), typeof(WriteValueFloat));
        }

        public void AutoClose() {
            FreeLibrary(netidx_excel_dll);
        }

        [ExcelFunction(Description = "Write data to netidx container", IsMacroType = false, IsExceptionSafe = false, IsThreadSafe = true, IsVolatile = false)]
        public static object NetSet(string path, object value)
        {
            try
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
                if (result == -1)
                    return "#SET";
                else
                    return (ExcelError)result;
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Description = "Subscribe to a netidx path", IsMacroType = false, IsExceptionSafe = false, IsThreadSafe = true, IsVolatile = false)]
        public static object NetGet(string path, object value)
        {
            return "TODO: IMPLEMENT ME";
        }
    }
}
