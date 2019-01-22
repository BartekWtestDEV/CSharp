using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using _Excel = Microsoft.Office.Interop.Excel;



namespace Tools
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelOperations.IsExcellInstalled();

            // ExcelOperations.CreateFile();

            // ExcelOperations.NewCreateFile(@"TEST");

             //ExcelOperations.AddUsersFromPreparedLisits();
            //ExcelOperations.AddSpecificUsers("Barbara","Wojnicz",27,7000,"27/08/2018","NOWY_SZID");

            //ExcelOperations.ClearSheetInExcel(2);
           ExcelOperations.CreateNewRandomPersons(100);


           // Console.ReadKey();

        }
    }


    public class NameSth
    {
        public string SheetName { get; set; }

    }
    public  static class ExcelOperations
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
                
                FileInfo excelFile = new FileInfo(@"C:\Users\wojnarowski_b\Desktop\trial\CreatedExcel.xlsx");
                excel.SaveAs(excelFile);
            }

        }

        public static void AddHeadRow(string FileName , string ID , string name , string surname,string birthTime, string slary )
        {
            ExcelPackage excel = new ExcelPackage();


        }

        public static void NewCreateFile(string path)
        {
            _Application excel = new _Excel.Application();
           
            Worksheet ws;
            Workbook wb =  excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            
            wb.SaveAs(@"C:\Users\wojnarowski_b\Desktop\trial\"+path);
            wb.Close();
        }
        public static void AddUsersFromPreparedLisits()
        {
            string[] names = {"Basia", "Arek", "Jan", "Zbigniew", "Magda" };
            string[] SURnames = { "KOW", "ZAD", "BAD", "KAT", "KOT" };
            int[] ages = { 22, 34, 24, 18, 50 };
            int[] Salaries = {1200,2900, 8900, 14000, 7000 };
            string[] startJobData = { "11/12/2018", "24/05/2007", "12/01/2001", "12/01/1991", "CPPP" };
            _Application excel = new _Excel.Application();
            Workbook sheet = excel.Workbooks.Open(@"C:\Users\wojnarowski_b\Desktop\trial\TEST.xlsx");
            Worksheet x = excel.ActiveSheet as Worksheet;
            for (int i = 0; i <= names.Count()-1; i++)
            {
                x.Cells[i+1, 1] = names[i];
                x.Cells[i+1, 2] = SURnames[i];
                x.Cells[i+1, 3] = ages[i];
                x.Cells[i+1, 4] = Salaries[i];
                x.Cells[i+1, 5] = startJobData[i];
               
            }

            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            


        }
        public static void AddSpecificUsers( string name , string surname , int age , int salary , string startJobData , string sheetName )
        {
            string Name = name;
            string Surname = surname;
            int Age = age;
            int Salary = salary;
            string StartJobData = startJobData; 

            _Application excel = new _Excel.Application();
            Workbook sheet = excel.Workbooks.Open(@"C:\Users\wojnarowski_b\Desktop\trial\TEST.xlsx");
           
            Worksheet x = excel.ActiveSheet as Worksheet;

            NameSth NameSheet = new NameSth();
            NameSheet.SheetName = sheetName;

            

            x.Name = NameSheet.SheetName;
            Microsoft.Office.Interop.Excel.Range userRange = x.UsedRange;
            int CountRecodrds = userRange.Rows.Count;
            int add = CountRecodrds + 1;


           
                x.Cells[add, 1] = Name;
                x.Cells[add, 2] = Surname;
                x.Cells[add, 3] = Age;
                x.Cells[add, 4] = Salary;
                x.Cells[add, 5] = StartJobData;

            

            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();



        }

        public static void ClearSheetInExcel(int sheetNumber)
        {


            _Application excel = new _Excel.Application();
            Workbook sheet = excel.Workbooks.Open(@"C:\Users\wojnarowski_b\Desktop\trial\TEST.xlsx");
            Worksheet ws = sheet.Worksheets[sheetNumber];
            ws.Cells.Clear();
            


            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();

        }

        public static void CreateNewRandomPersons(int HowMany)
        {

            _Application excel = new _Excel.Application();
            Workbook sheet = excel.Workbooks.Open(@"C:\Users\wojnarowski_b\Desktop\trial\TEST.xlsx");

            Worksheet x = excel.ActiveSheet as Worksheet;

            
            

            string[] names = { "Basia", "Arek", "Jan", "Zbigniew", "Magda", "Aga", "Michał", "Jan", "Kacper", "Andrzej", "Kazek" , "Bartek" , "Hipcio"  };
            string[] surnames = { "A", "B", "C", "D", "E", "F", "G", "H", "KOW", "ZAD", "BAD", "KAT", "KOT" };
            string[] startJobData = { "11/12/2018", "24/05/2007", "12/01/2001", "12/01/1991", "12/06/2017", "11/11/2018", "20/05/2007", "12/01/2000", "12/01/1990", "12/04/2007", "11/12/2018", "24/05/2007", "12/01/2002", "12/02/1991", "22/06/2017", "11/12/2018", "20/05/2002", "12/01/2002", "12/01/1992", "12/04/2007", };
            Random rnd = new Random();

            for (int k = 0; k <= HowMany; k++)
            {

             


          


                
                x.Cells[k+1, 1] = names[rnd.Next(1, names.Count())];
                x.Cells[k + 1, 2] = surnames[rnd.Next(1, surnames.Count())];
                x.Cells[k + 1, 3] = rnd.Next(18, 65);
                x.Cells[k + 1, 4] = rnd.Next(5000,20000);
                x.Cells[k + 1, 5] = startJobData[rnd.Next(1, startJobData.Count())];

                

            }



            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();



        }


    }

}
