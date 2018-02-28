using System;
using System.Text;
using OfficeOpenXml;
using System.IO;

public class libExcel_Epplus
{
	public Excel_Create(string SheetName)
	{
        ExcelPackage ExcelPkg = new ExcelPackage();
        ExcelWorksheet wsSheet1 = ExcelPkg.Workbook.Worksheets.Add("Sheet1");

        using (ExcelRange Rng = wsSheet1.Cells[2, 2, 2, 2])
        {
            Rng.Value = "Testing Excel from 4th Source";
            Rng.Style.Font.Size = 16;
            Rng.Style.Font.Bold = true;
            Rng.Style.Font.Italic = true;
        }
        wsSheet1.Protection.IsProtected = false;
        wsSheet1.Protection.AllowSelectLockedCells = false;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        ExcelPkg.SaveAs(new FileInfo(@"D:\New.xlsx"));
    }
}
