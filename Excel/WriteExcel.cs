using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace Excel
{
    class WriteExcel
    {
        public void WriteRange(DataTable dt, string path, string sheetName = "Sheet1", string startCell = "", bool addHeaders = false)
        {
            IXLWorkbook book;
            IXLWorksheet worksheet;

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

            if (addHeaders)
            {
                if (string.IsNullOrWhiteSpace(startCell))
                    startCell = GetCurentCell(worksheet);

                var indPosition = Convert.ToInt32(Regex.Match(startCell, @"\d+").Value);
                var cellAlfb = Regex.Match(startCell, @"[A-Z]+").Value;
                int number = 0;

                foreach (var b in cellAlfb)
                {
                    number += (int)b;
                }

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Console.WriteLine(dt.Columns[i]);
                    worksheet.Cell((number + i) > 90 ? ((char)Math.Floor(64 + (Convert.ToDouble(number) + i) / 90)).ToString() + ((char)(i + number % 90)).ToString() :
                        ((char)(number + i)).ToString() + indPosition).Value = dt.Columns[i];
                }

                startCell = cellAlfb + (indPosition + 1);
            }

            if (string.IsNullOrWhiteSpace(startCell))
                worksheet.Cell(GetCurentCell(worksheet)).Value = dt;
            else
                worksheet.Cell(startCell).Value = dt;

            book.SaveAs(path);
        }

        private string GetCurentCell(IXLWorksheet worksheet)
        {
            return Regex.Replace(worksheet.RangeUsed().LastRowUsed().FirstCell().ToString(), @"\d+",
                (Convert.ToInt32(Regex.Match(worksheet.RangeUsed().LastRowUsed().FirstCell().ToString(), @"\d+").Value) + 1).ToString());
        }
    }
}
