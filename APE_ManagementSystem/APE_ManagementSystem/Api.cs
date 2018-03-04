using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;
using MongoDB.Driver.GridFS;
using MongoDB.Driver.Linq;
using System.Collections;

using Excel = Microsoft.Office.Interop.Excel; 

namespace APE_ManagementSystem
{
   
    public class User
    {


        public int empNo { get; set; } //1

        public String name { get; set; } //4 
        public String date { get; set; } //6
        public Boolean absent { get; set; } //16
        public String overTime { get; set; } //
        public String workTime { get; set; } //

        


    }
    class Api
    {

        MongoDatabase db = DbConnection.conn();
        
        public void createData()
        {
            MongoCollection symbolcollection = db.GetCollection<Symbol>("Profile");
            

           // symbol.Name = "Star";
           // symbolcollection.Insert(symbol); 
           
        
            ArrayList al = new ArrayList(); // 1 4 6 10 11 16
            
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Syed Mustafa\Desktop\excel\test.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;


            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int ans = 0; int val = 0;
           
            for (int i = 2; i <= rowCount; i++)
            {
                User user = new User();




                try
                {
                    user.empNo = (int)xlRange.Cells[i, 1].Value;
                    user.name = "" + xlRange.Cells[i, 4].Value2;
                    user.date = "" + DateTime.Parse("" + xlRange.Cells[i, 6].Value.ToString());
                    user.absent = xlRange.Cells[i, 16].Value2;
                    user.overTime = "" + (((float)xlRange.Cells[i, 17].Value) * 24);
                    user.workTime = "" + (((float)xlRange.Cells[i, 18].Value) * 24);
                    symbolcollection.Insert(user);
                }
                catch(Exception e){
                
                }
                    
                
              /*  Console.WriteLine(xlRange.Cells[i, 1].Value2);
                Console.WriteLine(xlRange.Cells[i, 4].Value2);
                Console.WriteLine(xlRange.Cells[i, 6].Value);
                Console.WriteLine(xlRange.Cells[i, 17].Value2);
                Console.WriteLine(xlRange.Cells[i, 18].Value2);
                Console.WriteLine(xlRange.Cells[i, 16].Value2);
                Console.ReadLine(); */
            }
          //  String a = xlRange.Cells[1, 10].Value0;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //close and release
            xlWorkbook.Close();
            //quit and release
            xlApp.Quit();
        }
    }
}
