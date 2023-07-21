using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace netidxExcelDnaAddin
{
    public class Class1
    {
        // change your own dll directory
        [DllImport("J:\\office\\user\\ttang\\readonly\\netidx\\addin\\release\\lib\\netidx_excel.dll")]
        public static extern string write_value_string(string path, string value);

        [DllImport("J:\\office\\user\\ttang\\readonly\\netidx\\addin\\release\\lib\\netidx_excel.dll")]
        public static extern string write_value_int(string path, int value);

        [DllImport("J:\\office\\user\\ttang\\readonly\\netidx\\addin\\release\\lib\\netidx_excel.dll")]
        public static extern string write_value_float(string path, double value);


        [ExcelFunction(Description = "Write data to netidx container")]
        public static string PublishData(string path, object value)
        {
            if (value is int)
            {
                return write_value_int(path, (int)value);
            }
            else if (value is double || value is float)
            {
                return write_value_float(path, (double)value);
            }
            else if (value is string)
            {
                return write_value_string(path, (string)value);
            }
            else
            {
                return "input value is not an int, double or string";
            }
        }
    }
}
