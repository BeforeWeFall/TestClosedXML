using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Excel
{
    class ReadExcel
    {
        readonly DataTable dt = new DataTable();
        public  DataTable ReadSheet(string path, string sheetName, bool WithHeader = true, string Range="")
        {
            var exBook = new XLWorkbook(path).Worksheet(sheetName);

            if (string.IsNullOrEmpty(Range))
                return ReadAll(exBook, WithHeader);
            else
                return ReadCell(exBook, WithHeader,Range);
        }

        private DataTable ReadAll(IXLWorksheet book, bool WithHeader)
        {
            if (WithHeader)
                return book.RangeUsed().AsTable().AsNativeDataTable();
            else
            {
                CreateStandartHeaders(dt, book.RangeUsed().FirstRowUsed().CellCount());

                AddRange(dt, book.RangeUsed().Rows());

                return dt;
            }
        }
        
        private void CreateStandartHeaders(DataTable dt, int count)
        {
            for (int i = 0; i < count ; i++)
            {
                dt.Columns.Add(GetAlfb(i) , typeof(String));
            }
        }

        private void AddRange(DataTable dt, IXLRangeRows Range)
        {
            foreach (var row in Range)
            {
                dt.Rows.Add();
                int i = 0;
                foreach (IXLCell cell in row.Cells())
                {
                    string val = string.Empty;
                    try
                    {
                        val = cell.Value.ToString();
                    }
                    catch { }

                    dt.Rows[dt.Rows.Count - 1][i] = val;
                    i++;
                }
            }
        }

        private string GetAlfb(int num)
        {
            return (065 + num) > 90 ? ((char)Math.Floor(64 + (64.0 + num) / 90)).ToString() + ((char)( num % 90)).ToString() : ((char)(065 + num)).ToString();
        }

        private DataTable ReadCell(IXLWorksheet book, bool WithHeader, string Range)
        {

            var bookRange = Range.Contains(":") ? book.Range(Range.Split(":")[0].ToUpper(), Range.Split(":")[1].ToUpper()) :
                book.Range(Range.ToUpper(), GetAlfb(book.RangeUsed().FirstRowUsed().CellCount()-1) + (book.RangeUsed().RowCount()));

            if (WithHeader)
            {
                return bookRange.AsTable().AsNativeDataTable();
            }         
            else 
            {
                CreateStandartHeaders(dt, bookRange.FirstRowUsed().CellCount());

                AddRange(dt, bookRange.Rows());

                return dt;
            }
        }
    }
}
