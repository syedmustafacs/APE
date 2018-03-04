using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;  

namespace APE_ManagementSystem
{
    class ExcelProcessor
    {
        public void getFile() {
            ArrayList al = new ArrayList(); // 1 4 6 10 11 16
            User user = new User();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Syed Mustafa\Desktop\testdata.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int ans = 0; int val = 0;

            for (int i = 1; i <= rowCount; i++)
            {
                al.Add( (int)xlRange.Cells[i, 4].Value2.ToString());
                al.Add( xlRange.Cells[i, 6].Value2.ToString());
                 al.Add(xlRange.Cells[i, 10].Value2.ToString());
                 al.Add(xlRange.Cells[i, 11].Value2.ToString());
                 al.Add( xlRange.Cells[i, 16].Value2.ToString());
                
            }
           

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //close and release
            xlWorkbook.Close();
     //quit and release
            xlApp.Quit();
        }
        public void getExcelFile()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Syed Mustafa\Desktop\testdata.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
           

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int ans = 0; int val = 0;

            for (int i = 1; i <= rowCount; i++)
            {
                String a = xlRange.Cells[i, 9].Value2.ToString();
                val = int.Parse(a);
                ans = ans + val;
            }
            Console.WriteLine("Here comess  " + ans);
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                }
            }


            GC.Collect();
            GC.WaitForPendingFinalizers();



            //close and release
            xlWorkbook.Close();


            //quit and release
            xlApp.Quit();

        }
    }
}
