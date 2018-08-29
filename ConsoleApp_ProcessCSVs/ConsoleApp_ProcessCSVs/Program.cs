using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApp_ProcessCSVs
{
    class Program
    {
        static void Main(string[] args)
        {
            foreach (var file in Directory.GetFiles(@"T:\\ARS-DataMgmt\\laura.j.ogan\\NUOnet_csv\\", "*.csv"))
            {
                Application xlApp = new Application();
                Workbook xlWorkbook = xlApp.Workbooks.Open(file, Type.Missing, true);
                _Worksheet xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];

                Console.Write(xlWorksheet.Name + " fld2Base = [\"");

                Range xlRange = xlWorksheet.UsedRange;
                foreach (Range c in xlRange.Rows.Cells)
                {
                    Console.Write(c.Value2 + "\",\"");
                }

                Console.Write("]");
                Console.WriteLine();

                Console.Write(xlWorksheet.Name + " fld2New  = [\"");

                foreach (Range c in xlRange.Rows.Cells)
                {
                    Console.Write(c.Value2 + "_1" + "\",\"");
                }

                Console.Write("]");
                Console.WriteLine();

                xlWorkbook.Close();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);

            }
            Console.ReadKey();
        }
    }
}
