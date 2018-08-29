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
    // The purpose of this script is to:
    //      loop through all .csv files in a folder, 
    //      transpose all the cells in the range (in this case CloumnA), 
    //      put quotes around the values, 
    //      add begining and ending wrappers including the worksheet name, 
    //      add a second line that is unique from the first yet using the same source values.

    {
        static void Main(string[] args)
        {
            // Set the working directory containing the .csv files
            foreach (var file in Directory.GetFiles(@"C:\\C_Working\\", "*.csv"))
            {
                // Grab the file
                Application xlApp = new Application();
                Workbook xlWorkbook = xlApp.Workbooks.Open(file, Type.Missing, true);
                _Worksheet xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];

                // Clear blank cell from range before setting range
                xlWorksheet.Columns.ClearFormats();
                xlWorksheet.Rows.ClearFormats();

                // Declare the range of cells
                Range xlRange = xlWorksheet.UsedRange;
                Range last = xlWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                // Make a string of names from the range of data
                Console.Write(xlWorksheet.Name + " fld2Base = [\"");

                foreach (Range c in xlRange.Rows.Cells)
                {
                    if (c.Value2 != last.Value2)
                        Console.Write(c.Value2 + "\",\"");
                }
                Console.Write(last.Value2 + "\"]");
                Console.WriteLine();

                // Make a unique string of names by adding the "_1"
                Console.Write(xlWorksheet.Name + " fld2New  = [\"");

                foreach (Range c in xlRange.Rows.Cells)
                {
                    if (c.Value2 != last.Value2)
                        Console.Write(c.Value2 + "_1\",\"");
                }

                Console.Write(last.Value2 + "_1\"]");
                Console.WriteLine();

                // Finish up, and don't save changes to file
                xlWorkbook.Close(false);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);

            }
            // The script is set to keep the console open until you copy your text and manually close the window
            Console.ReadKey();
        }
    }
}
