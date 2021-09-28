using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\Users\ULBelykh\Downloads\Telegram Desktop\Список слов.xlsx";
            string listName = "Лист1";
            var dt = new DataTable();
            //Border er = new Border(path, listName, "A17:A19");

            //Console.WriteLine(Convert.ToDecimal("88.17").ToString());
            //Console.WriteLine(Convert.ToDecimal("88,17").ToString());
            //ColorExcel er = new ColorExcel(path, listName,"A2");
            SetTip er = new SetTip(path, listName, "A17:A19", 2);
            //for (int i = 0; i < 100; i++)
            //{
            //    er = new SetTip(path, listName, "A"+(1+i), i);
            //}


            //dt = new ReadExcel().ReadSheet(path, listName, false); //последние два можно не указывать

            //foreach (DataColumn e in dt.Columns)
            //    Console.Write(e.ColumnName + "_");
            //Console.WriteLine("");
            //Console.WriteLine("Zagolovki");

            //foreach (DataRow t in dt.Rows)
            //{
            //    foreach (DataColumn e in dt.Columns)
            //        Console.Write(t[e] + "+");
            //    Console.WriteLine("");
            //}
            //Console.WriteLine("");
            //.Style.NumberFormat.NumberFormatId = 1;
            //new WriteExcel().WriteRange(dt, @"C:\Users\ULBelykh\Desktop\Отчет6.xlsx", "Sheet1", "G2", true); //последние два можно не указывать

        }


    }
}
