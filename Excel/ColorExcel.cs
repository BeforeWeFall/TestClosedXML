using ClosedXML.Excel;

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Excel
{
    class ColorExcel
    {
        private IXLWorkbook book;
        private IXLWorksheet worksheet;
        public ColorExcel(string path, string sheetName, string target)
        {
            if (File.Exists(path))
            {
                book = new XLWorkbook(path);

                try
                {
                    worksheet = book.Worksheet(sheetName);
                }
                catch
                {
                    worksheet = book.AddWorksheet(sheetName);
                }
            }
            else
            {
                book = new XLWorkbook();
                worksheet = book.AddWorksheet(sheetName);
            }

            if (!target.Contains(":"))
                ColorCell(target);
            else
                ColorRange(target);
            book.Save();
            //wb.SaveAs(filePath);
        }
        
        private void ColorCell(string target) 
        {   
            worksheet.Cell(target).Style.Fill.BackgroundColor = XLColor.FromColor(System.Drawing.Color.Red);
        }
        private void ColorRange(string target) 
        {
            string[] range = target.Split(":");
            
            IXLRange e;
            if (string.IsNullOrWhiteSpace(range[1]))
            {
                e=worksheet.Range(range[0].ToUpper(), GetAlfb(worksheet.RangeUsed().FirstRowUsed().CellCount() - 1) + (worksheet.RangeUsed().RowCount()));
            }
            else
            {
               e = worksheet.Range(range[0].ToUpper(), range[1].ToUpper());
            }
            
            e.Style.Fill.BackgroundColor = XLColor.Green;
        }
        private string GetAlfb(int num)
        {
            return (065 + num) > 90 ? ((char)Math.Floor(64 + (64.0 + num) / 90)).ToString() + ((char)(num % 90)).ToString() : ((char)(065 + num)).ToString();
        }
    }
}
