using System;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;


namespace PruebaExcel_EPplus
{
    class Program
    {
        static void Main(string[] args)
        {
            libExcel_epp l = new libExcel_epp();

            l.Excel_Create("Sheet1",01,02,"Test1","www.google.com");
            //l.Excel_Create("Sheet1", 02, 03, "Test2", "www.google.com");
            //l.Excel_Create("Sheet1", 04, 05, "Test3", "www.google.com");
            l.Excel_Mod(4,1,"Hello");
            l.Excel_Mod(3, 10, 10, "10", "www.google.com");
            //l.Excel_Get();

            //ExcelPackage package = new ExcelPackage(new FileInfo(@"D:\New2.xlsx"));
            //ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
            //IEnumerable<libExcel_epp> records = QuerysRecord(workSheet, true);
            //Console.WriteLine("Records count: " + records.Count());
            //Console.ReadLine();

        }

        /// <summary>
        /// Populate award objects from spreadsheet
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="firstRowHeader"></param>
        /// <returns></returns>
        /// 
        static IEnumerable<libExcel_epp> QuerysRecord(ExcelWorksheet workSheet, bool firstRowHeader)
        {
            IList<libExcel_epp> records = new List<libExcel_epp>();

            if (workSheet != null)
            {
                Dictionary<string, int> header = new Dictionary<string, int>();

                for (int rowIndex = workSheet.Dimension.Start.Row; rowIndex <= workSheet.Dimension.End.Row; rowIndex++)
                {
                    //Assume the first row is the header.  Then use the column match ups by name to determine the index.
                    //This will allow you to have the order of the columns change without any affect.

                    if (rowIndex == 1 && firstRowHeader)
                    {
                        header = ExcelHelper.GetExcelHeader(workSheet, rowIndex);
                    }
                    else
                    {
                        records.Add(new libExcel_epp
                        {
                            ID = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "ID"),
                            Num = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Num"),
                            String = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "String")
                        });

                    }
                }
            }

            return records;
        }
    }
}
