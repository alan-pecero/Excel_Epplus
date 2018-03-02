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
            libExcel_epp objectoExcel = new libExcel_epp();

            //libExcel_epp objeto_2 = new libExcel_epp("Worksheet1");
            //libExcel_epp objeto_3 = new libExcel_epp("Worksheet2","Sheet 3");
            //objectoExcel.WorkbookCreate("Workbook3");
            //objectoExcel.WorkbookCreate(@"E:\","WorkbookPrueba","Sheet 40");
            objectoExcel.WorkbookCreate("Workbook", "Sheet 1");
            //objectoExcel.WorkbookDelete("Workbook2");
            //objectoExcel.WorkbookDelete(@"E:\","WorkbookPrueba");
            //libExcel_epp objeto_4 = new libExcel_epp(@"E:\trabajo\", "Workbook50", "Worksheet100" );
            //objectoExcel.WorkbookOverwrite("WorkbookPrueba")
            Console.WriteLine("El archivo creado se encuentra en: " + objectoExcel.m_fileInfo.DirectoryName);
            Console.ReadKey();

            /**
            objectoExcel.WorksheetCreate("Jola");
            objectoExcel.WorksheetStyling("Jola",14,"Times New Roman");
            objectoExcel.WorksheetCreate("Hoja 3");
            objectoExcel.WorksheetStyling(@"E:\","WorkbookPrueba","Hoja 3",14,"Times New Roman");
            objectoExcel.WorksheetStyling("WorkbookPrueba","Sheet 40",14,"Times New Roman"); //Si le doy un workbook diferente con el objeto que tengo ya creado, entonces lo sobreescribe y ya
            */



            //objectoExcel.WorkbookDelete(@"E:\","WorkbookPrueba");
            //Console.WriteLine(objectoExcel.m_fileInfo.ToString());
            //Console.ReadKey();


            objectoExcel.Excel_Create("Sheet1", 01, 02, "Test1", "www.google.com");
            objectoExcel.Excel_Mod(6, 1, "Hola");
            objectoExcel.Excel_Mod(3, 10, 10, "10", "www.google.com");
            objectoExcel.Excel_Mod(4, 11, 11, "11", "www.google.com");
            objectoExcel.Excel_Mod(5, 12, 12, "12", "www.google.com");

            ExcelPackage package = new ExcelPackage(new FileInfo(@"D:\FinalTest.xlsx"));
            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
            IEnumerable<libExcel_epp> records = QuerysRecord(workSheet, true);
            Console.WriteLine("Records count: " + records.Count());
            Console.ReadLine();
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
                        header = libExcel_epp.GetExcelHeader(workSheet, rowIndex);
                    }
                    else
                    {
                        records.Add(new libExcel_epp
                        {
                            ID = libExcel_epp.ParseWorksheetValue(workSheet, header, rowIndex, "ID"),
                            Num = libExcel_epp.ParseWorksheetValue(workSheet, header, rowIndex, "Num"),
                            String = libExcel_epp.ParseWorksheetValue(workSheet, header, rowIndex, "String")
                        });

                    }
                }
            }

            return records;
        }
    }
}
