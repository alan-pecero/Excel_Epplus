using System;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Xml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System.Linq;
using System.Data;
using System.Threading.Tasks;

namespace PruebaExcel_EPplus
{

    class libExcel_epp
    {
        //atributes
        public ExcelPackage m_objExcel; //the Excel app itself
        public ExcelWorkbook m_objWorkbook; //A workbook object we will be using constantly
        public ExcelWorksheet m_objWorksheet; //A worksheet object because we need to create at least one worksheet
        private string m_strFileName; //the name of the file we are using, this is used for Workbooks as those are the files themselves
        public FileInfo m_fileInfo; //the path of the directory where we want to save the file
        private string m_strFontName; //the name of the font we want to use
        private int m_intFontSize; //the size of the font we want to use
        Dictionary<KeyValuePair<int, int>, object> values;
        bool m_bTitle = false;


        //properties
        public FileInfo ProprFileInfo { get => m_fileInfo; set => m_fileInfo = value; }
        private string ProprFileName { get => m_strFileName; set => m_strFileName = value; }
        private string ProprFontName { get => m_strFontName; set => m_strFontName = value; }
        private int ProprFontSize { get => m_intFontSize; set => m_intFontSize = value; }
        public string ID { get; set; }
        public string Num { get; set; }
        public string String { get; set; }

        //methods
        public libExcel_epp() // default constructor
        {
            this.m_objExcel = new ExcelPackage();
        }

        public libExcel_epp(string pstrWorkbookName) // Constructor that receives the name of the Woorkbook as it's parameter
        {
            this.m_objExcel = new ExcelPackage();
            this.m_strFileName = pstrWorkbookName;
            WorkbookCreate(m_strFileName);
            m_objExcel.Save();
        }

        public libExcel_epp(string pstrWorkbookName, string pstrWorksheetName) // Constructor that receives the name of the Workbook and a Worksheet to start
        {
            this.m_objExcel = new ExcelPackage();
            this.m_strFileName = pstrWorkbookName;
            WorkbookCreate(m_strFileName, pstrWorksheetName);
            m_objExcel.Save();
        }

        public void WorkbookCreate() //creates a Workbook with a default name for both the file and the first worksheet
        {
            /**
            var newFile = new FileInfo(outputDirectory.FullName + @"\Workbook.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDirectory.FullName + @"\Workbook.xlsx");
            }

            using (var package = new ExcelPackage(newFile));
            */
            m_objExcel.Workbook.Worksheets.Add("Sheet 1");
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); //this line is used to encode the file so that it can be saved where we want to
            m_objExcel.SaveAs(m_fileInfo = new FileInfo(@"E:\Workbook.xlsx"));
            m_objExcel.Save();
        }

        public void WorkbookCreate(string pstrWorkbookName, string pstrWorksheetName) // creates a new Workbook with both the name of the Workbook and the name of the first worksheet
        {
            FileInfo file = new FileInfo(@"E:\" + pstrWorkbookName + ".xlsx");
            if (file.Exists)
            {
                Console.WriteLine("Existing archive. Overwriting");
                OverwriteExcel(pstrWorkbookName);
            }
            else
            {
                m_objExcel.Workbook.Worksheets.Add(pstrWorksheetName);
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                m_objExcel.SaveAs(m_fileInfo = new FileInfo(@"E:\" + pstrWorkbookName + ".xlsx"));
                m_objExcel.Save();
            }
           
        }

        public void WorkbookCreate(string pstrWorkbookName) //creates a new Workbook with a specified name, the first worksheet is named "Sheet 1"
        {
            m_objExcel.Workbook.Worksheets.Add("Sheet 1");
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            m_objExcel.SaveAs(m_fileInfo = new FileInfo(@"E:\" + pstrWorkbookName + ".xlsx"));
            m_objExcel.Save();
        }

        public void WorkbookCreate(string pstrRoute, string pstrWorkbookName, string pstrWorksheetName) //creates a new Workbook with a specified route, Workbook name and worksheet name
        {
            m_objExcel.Workbook.Worksheets.Add(pstrWorksheetName);
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            m_objExcel.SaveAs(m_fileInfo = new FileInfo(pstrRoute + pstrWorkbookName + ".xlsx")); //for now, it only saves to an already existing directory
            m_objExcel.Save();
        }

        public void WorkbookDelete(string pstrWorkbookName) //deletes a Workbook with a specified name
        {
            FileInfo file = new FileInfo(@"E:\" + pstrWorkbookName + ".xlsx");
            if (file.Exists)
                file.Delete();
        }

        public void WorkbookDelete(string pstrRoute, string pstrWorkbookName) //deletes a Workbook in a specified route with a specified name
        {
            FileInfo file = new FileInfo(pstrRoute + pstrWorkbookName + ".xlsx");

            if (file.Exists)
                file.Delete();
        }


        //Creates a Worksheet for an existing Workbook (for some reason it can only do that if the Workbook is immediatly created before this is called. Weird)
        //update 2/26/2018 Now I know why: because you have to create the object first so you can manipulate it at run-time
        public void WorksheetCreate(string pstrWorksheetName)
        {
            ExcelWorksheet worksheet = m_objExcel.Workbook.Worksheets.Add(pstrWorksheetName);
            m_strFileName = m_fileInfo.FullName;
            Console.WriteLine("Agregando una nueva worksheet a: " + m_strFileName);
            Console.ReadKey();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            m_objExcel.SaveAs(m_fileInfo);
            m_objExcel.Save();
        }

        //This will be here for now 
        public void OverwriteExcel(string pstrFileName)
        {
            m_objExcel.File.Replace(pstrFileName, pstrFileName);
            m_objExcel.Save();
        }

        //Added the ability to open a file from a FileStream and modify it, so I did what I wanted to do for a while which was edit an existing file
        //this should take the name of a Worksheet and at some styling and format to it. Either or both.+, I'll check what I can do
        public void WorksheetStyling(string pstrWorksheetName, int pintFontSize, string pstrFontName)
        {
            using (FileStream stream = new FileStream(@"E:\WorkbookPrueba.xlsx", FileMode.Open))
            {
                m_objExcel.Load(stream);
                ExcelWorksheet worksheet = m_objExcel.Workbook.Worksheets[pstrWorksheetName];
                worksheet.TabColor = Color.Blue;
                worksheet.DefaultRowHeight = 12;

                using (var range = worksheet.Cells[1, 1, 1, 5])
                {
                    range.Style.Font.Size = pintFontSize;
                    range.Style.Font.Name = pstrFontName;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Black);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.ShrinkToFit = false;
                }

                using (var range = worksheet.Cells[2, 1, 5, 5])
                {
                    range.Style.Font.Size = pintFontSize;
                    range.Style.Font.Name = pstrFontName;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    range.Style.Font.Color.SetColor(Color.WhiteSmoke);
                    range.Style.ShrinkToFit = false;
                }
            }
            m_objExcel.Save();
        }

        public void WorksheetStyling(string pstrWorkbookName, string pstrWorksheetName, int pintFontSize, string pstrFontName)
        {
            using (FileStream stream = new FileStream(@"E:\"+ pstrWorkbookName +".xlsx", FileMode.Open))
            {
                m_objExcel.Load(stream);
                ExcelWorksheet worksheet = m_objExcel.Workbook.Worksheets[pstrWorksheetName];

                worksheet.DefaultRowHeight = 12;
                worksheet.TabColor = Color.OrangeRed;

                using (var range = worksheet.Cells[1, 1, 1, 5])
                {
                    range.Style.Font.Size = pintFontSize;
                    range.Style.Font.Name = pstrFontName;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Black);
                    range.Style.Font.Color.SetColor(Color.White);
                    range.Style.ShrinkToFit = false;
                }

                using (var range = worksheet.Cells[2, 1, 5, 5])
                {
                    range.Style.Font.Size = pintFontSize;
                    range.Style.Font.Name = pstrFontName;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    range.Style.Font.Color.SetColor(Color.BlueViolet);
                }
            }
            m_objExcel.Save();
        }

        public void WorksheetStyling(string pstrRoute, string pstrWorkbookName, string pstrWorksheetName, int pintFontSize, string pstrFontName)
        {
            using (FileStream stream = new FileStream(pstrRoute + pstrWorkbookName + ".xlsx", FileMode.Open))
            {
                m_objExcel.Load(stream);
                ExcelWorksheet worksheet = m_objExcel.Workbook.Worksheets[pstrWorksheetName];

                worksheet.DefaultRowHeight = 12;
                worksheet.TabColor = Color.OrangeRed;

                using (var range = worksheet.Cells[1, 1, 1, 5])
                {
                    range.Style.Font.Size = pintFontSize;
                    range.Style.Font.Name = pstrFontName;
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    range.Style.Font.Color.SetColor(Color.Black);
                    range.Style.ShrinkToFit = false;
                }

                using (var range = worksheet.Cells[2, 1, 5, 5])
                {
                    range.Style.Font.Size = pintFontSize;
                    range.Style.Font.Name = pstrFontName;
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
                    range.Style.Font.Color.SetColor(Color.Violet);
                }
            }
            m_objExcel.Save();
        }


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
            m_bTitle = true;
            var row = datatable.NewRow();
            row[0] = col1;
            row[1] = col2;
            row[2] = col3 + " " + Path.GetRandomFileName();
            datatable.Rows.Add(row);

            //Create a test file
            var existingFile = new FileInfo(@"D:\FinalTest.xlsx");
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

        //This Excel_Mod method it's just for update a single cell  
        public void Excel_Mod(int row, int col, string val)
        {
            FileInfo file = new FileInfo(@"D:\FinalTest.xlsx");
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

        //This Excel_Mod method it's just for update or add a entire row 
        public void Excel_Mod(int row, int col1, int col2, string col3, string URL)
        {
            FileInfo file = new FileInfo(@"D:\FinalTest.xlsx");
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

        

        /// <summary>
        /// Gets the excel header and creates a dictionary object based on column name in order to get the index.
        /// Assumes that each column name is unique.
        /// </summary>
        /// <param name="workSheet"></param>
        /// <returns></returns>
        public static Dictionary<string, int> GetExcelHeader(ExcelWorksheet workSheet, int rowIndex)
        {
            Dictionary<string, int> header = new Dictionary<string, int>();

            if (workSheet != null)
            {
                for (int columnIndex = workSheet.Dimension.Start.Column; columnIndex <= workSheet.Dimension.End.Column; columnIndex++)
                {
                    if (workSheet.Cells[rowIndex, columnIndex].Value != null)
                    {
                        string columnName = workSheet.Cells[rowIndex, columnIndex].Value.ToString();

                        if (!header.ContainsKey(columnName) && !string.IsNullOrEmpty(columnName))
                        {
                            header.Add(columnName, columnIndex);
                        }
                    }
                }
            }

            return header;
        }

        /// <summary>
        /// Parse worksheet values based on the information given.
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static string ParseWorksheetValue(ExcelWorksheet workSheet, Dictionary<string, int> header, int rowIndex, string columnName)
        {
            string value = string.Empty;
            int? columnIndex = header.ContainsKey(columnName) ? header[columnName] : (int?)null;

            if (workSheet != null && columnIndex != null && workSheet.Cells[rowIndex, columnIndex.Value].Value != null)
            {
                value = workSheet.Cells[rowIndex, columnIndex.Value].Value.ToString();
            }

            return value;
        }
    }
}
