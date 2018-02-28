using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Data;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System.Drawing;

namespace PruebaExcel_EPplus
{
    class libExcel_epp
    {
        //Dictionary<string, int> dictionary = new Dictionary<string, int>();
        Dictionary<KeyValuePair<int, int>, object> values;
        bool bTitle = false;

        public string ID { get; set; }
        public string Num { get; set; }
        public string String { get; set; }

        public void Excel_Create(string SheetName, int col1, int col2, string col3, string URL)
        {
            //ExcelPackage ExcelPkg = new ExcelPackage();
            //ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add(SheetName);

            //using (ExcelRange Rng = wsSheet1.Cells[row, col])
            //{
            //    Rng.Value = val;
            //    Rng.Style.Font.Size = 16;
            //    Rng.Style.Font.Bold = true;
            //    Rng.Style.Font.Italic = true;
            //}
            
            var datatable = new DataTable("tblData");
            //Generate titles of datatable
            datatable.Columns.AddRange(new[] { new DataColumn("ID", typeof(int)), new DataColumn("Num", typeof(int)), new DataColumn("String", typeof(object)), new DataColumn("Screenshot", typeof(string)) });
            bTitle = true;
            var row = datatable.NewRow();
            row[0] = col1;
            row[1] = col2;
            row[2] = col3 +" "+ Path.GetRandomFileName();
            datatable.Rows.Add(row);

            //Create a test file
            var existingFile = new FileInfo(@"D:\TestFinal.xlsx");
            if (existingFile.Exists)
                existingFile.Delete();

            using (var pck = new ExcelPackage(existingFile))
            {
                var worksheet = pck.Workbook.Worksheets.Add(SheetName);
                worksheet.Cells.LoadFromDataTable(datatable, true);
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                pck.Save();
            }

            using (var pck = new ExcelPackage(existingFile))
            {
                var worksheet = pck.Workbook.Worksheets[SheetName];

                //Cells only contains references to cells with actual data
                var cells = worksheet.Cells;
                var dictionary = cells.GroupBy(c => new { c.Start.Row, c.Start.Column })
                    .ToDictionary(
                        rcg => new KeyValuePair<int, int>(rcg.Key.Row, rcg.Key.Column),
                        rcg => cells[rcg.Key.Row, rcg.Key.Column].Value);

                foreach (var kvp in dictionary)
                {
                    Console.WriteLine("{{ Row: {0}, Column: {1}, Value: \"{2}\" }}", kvp.Key.Key, kvp.Key.Value, kvp.Value);
                }
                Console.ReadLine();


                string StyleName = "HyperStyle";
                ExcelNamedStyleXml HyperStyle = worksheet.Workbook.Styles.CreateNamedStyle(StyleName);
                HyperStyle.Style.Font.UnderLine = true;
                HyperStyle.Style.Font.Size = 12;
                HyperStyle.Style.Font.Color.SetColor(Color.Blue);

                //------HYPERLINK to a website.  
                using (ExcelRange Rng = worksheet.Cells[2, 4, 2, 4])
                {
                    Rng.Hyperlink = new Uri("http://" + URL, UriKind.Absolute);
                    Rng.Value = "Screenshot";
                    Rng.StyleName = StyleName;
                }


                worksheet.Protection.IsProtected = false;
                worksheet.Protection.AllowSelectLockedCells = false;
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                pck.SaveAs(new FileInfo(@"D:\TestFinal.xlsx"));
            }
        }

        //Excel_Mod is a method for update data from a file already created 
        public void Excel_Mod(int row, int col, string val)
        {
            FileInfo file = new FileInfo(@"D:\TestFinal.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets[1];
                excelWorksheet.Cells[row, col].Value = val;
                //excelWorksheet.Cells[3, 2].Value = "Test2";
                //excelWorksheet.Cells[3, 3].Value = "Test3";

                excelPackage.Save();
            }
        }
        //Excel_Mod is a method for update data from a file already created 
        public void Excel_Mod(int row, int col1, int col2, string col3, string URL)
        {
            FileInfo file = new FileInfo(@"D:\TestFinal.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets[1];
                excelWorksheet.Cells[row, 1].Value = col1;
                excelWorksheet.Cells[row, 2].Value = col2;
                excelWorksheet.Cells[row, 3].Value = col3;
                excelWorksheet.Cells[row, 4].Value = URL;

                //Update the field "Screenshot"
                string StyleName = "HyperStyle";
                //ExcelNamedStyleXml HyperStyle = excelWorksheet.Workbook.Styles.CreateNamedStyle(StyleName);
                //HyperStyle.Style.Font.UnderLine = true;
                //HyperStyle.Style.Font.Size = 12;
                //HyperStyle.Style.Font.Color.SetColor(Color.Blue);
                using (ExcelRange Rng = excelWorksheet.Cells[row, 4, row, 4])
                {
                    Rng.Hyperlink = new Uri("http://" + URL, UriKind.Absolute);
                    Rng.Value = "Screenshot";
                    Rng.StyleName = StyleName;
                }

                excelPackage.Save();
            }
        }

        //Test about getting data a set on a dictionary
        public void Excel_Get()
        {
            var datatable = new DataTable("tblData");
            //Generate titles of datatable
            datatable.Columns.AddRange(new[] { new DataColumn("ID", typeof(int)), new DataColumn("Num", typeof(int)), new DataColumn("String", typeof(object)) });
            bTitle = true;

            for (var i = 0; i < 10; i++)
            {
                var row = datatable.NewRow();
                row[0] = i;
                row[1] = i * 10;
                row[2] = Path.GetRandomFileName();
                datatable.Rows.Add(row);
            }

            //Create a test file
            var existingFile = new FileInfo(@"D:\New2.xlsx");
            if (existingFile.Exists)
                existingFile.Delete();

            using (var pck = new ExcelPackage(existingFile))
            {
                var worksheet = pck.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells.LoadFromDataTable(datatable, true);
                pck.Save();
            }

            using (var pck = new ExcelPackage(existingFile))
            {
                var worksheet = pck.Workbook.Worksheets["Sheet1"];

                //Cells only contains references to cells with actual data
                var cells = worksheet.Cells;
                var dictionary = cells.GroupBy(c => new { c.Start.Row, c.Start.Column })
                    .ToDictionary(
                        rcg => new KeyValuePair<int, int>(rcg.Key.Row, rcg.Key.Column),
                        rcg => cells[rcg.Key.Row, rcg.Key.Column].Value);

                foreach (var kvp in dictionary)
                {
                    Console.WriteLine("{{ Row: {0}, Column: {1}, Value: \"{2}\" }}", kvp.Key.Key, kvp.Key.Value, kvp.Value);
                }
                Console.ReadLine();
            }
        }
    }
}
