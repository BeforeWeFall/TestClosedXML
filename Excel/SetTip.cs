using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Excel
{
    class SetTip
    {
        private IXLWorkbook book;
        private IXLWorksheet worksheet;
        public SetTip(string path, string sheetName, string target, int idTip)
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
                TipCell(target, idTip);
            else
                TipRange(target, idTip);
            book.Save();
            //wb.SaveAs(filePath);
        }

        private void TipCell(string target, int idTip)
        {
            //worksheet.Cell(target).Value= Convert.ToDecimal(worksheet.Cell(target).Value);
            //worksheet.Cell(target).SetDataType(XLDataType.Number);
            //worksheet.Cell(target).Style.NumberFormat.NumberFormatId = idTip; //переделать на это

            //worksheet.Cell(target).DataType = XLDataType.Number;
            //worksheet.Cell(target).Style.NumberFormat.Format = "2";

            worksheet.Cell(target).Value = Convert.ToDecimal(worksheet.Cell(target).GetDouble());

                worksheet.Cell(target).SetDataType(XLDataType.Number);
            worksheet.Cell(target).Style.NumberFormat.NumberFormatId = idTip;

        }
        private void TipRange(string target, int idTip)
        {
            string[] range = target.Split(":");

            IXLRange rangeXL;
            if (string.IsNullOrWhiteSpace(range[1]))
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), GetAlfb(worksheet.RangeUsed().FirstRowUsed().CellCount() - 1) + (worksheet.RangeUsed().RowCount()));
                
            }
            else
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), range[1].ToUpper());
            }
            
            if (idTip > 11)
                rangeXL.SetDataType(XLDataType.DateTime);
            else if (idTip > 0)
            {
                foreach (var cell in rangeXL.Cells())
                {
                    cell.Value = Convert.ToDecimal(cell.Value);
                }
                rangeXL.SetDataType(XLDataType.Number);
            }
                
            rangeXL.Style.NumberFormat.NumberFormatId = idTip;
            
        }
        private string GetAlfb(int num)
        {
            return (065 + num) > 90 ? ((char)Math.Floor(64 + (64.0 + num) / 90)).ToString() + ((char)(num % 90)).ToString() : ((char)(065 + num)).ToString();
        }
    }
}
