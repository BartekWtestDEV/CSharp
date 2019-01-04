using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;


namespace Tools
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelOperations.IsExcellInstalled();

            ExcelOperations.CreateFile();


            Console.ReadKey();

        }
    }

    public static class ExcelOperations
    {


        public static void IsExcellInstalled()
        {
            var isExcelInstalled = Type.GetTypeFromProgID("Excel.Application");
            if (isExcelInstalled == null)
            {
                Console.WriteLine("NO");

            }
            else
            {
                Console.WriteLine("Yes");
            }



        }

        public static void CreateFile()
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");
                
                FileInfo excelFile = new FileInfo(@"C:\Users\Bartek\Desktop\CreatedExcel.xlsx");
                excel.SaveAs(excelFile);
            }

        }

        public static void AddHeadRow(string FileName , string ID , string name , string surname,string birthTime, string slary )
        {
            ExcelPackage excel = new ExcelPackage();
            excel.Workbook.Worksheets

        }



    }

}
