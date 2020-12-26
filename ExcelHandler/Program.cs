using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Data;
using System.IO;

namespace ExcelHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Program Start");
            string file = Console.ReadLine();
            file.Trim();
                
            Stopwatch timer = new Stopwatch();
            ExcelSheet sheet = null;
            try
            {
                sheet = new ExcelSheet(file);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }

            if (sheet == null)
            {
                Console.WriteLine("sheet cannot be read");
                AskKeyToExit();
                return;
            }
            

            Console.WriteLine($"Rows          : {sheet.RowsCount}");
            Console.WriteLine($"Colunms       : {sheet.ColumnsCount}");
            Console.WriteLine($"Not null data : {sheet.RowsCount}");

            Console.Write($"Column Names(Header) : ");
            foreach (var col in sheet.GetColumnNames())
            {
                Console.Write($"{col}, ");
            }

            for (int i = 1; i <= sheet.RowsCount; i++) 
            {
                if (i >= 10)
                {
                    Console.WriteLine("---------------------------------------sample 10 rows---------------------------------------");
                    break;
                }

                for (int j = 1; j <= sheet.ColumnsCount; j++)
                {
                    Console.WriteLine(sheet[i, j] + ":\t" + sheet[i, j].GetType());
                }
                Console.WriteLine("\n");
            }

            Console.WriteLine("Converting DataTable...");
            timer.Start();
            using(DataTable dt = sheet.GetFullDataTable_Sync())
            {               
            }
            timer.Stop();
            Console.WriteLine("...Done!");
            Console.WriteLine($"time elapsed to convert DataTable : {timer.ElapsedMilliseconds} ms (Rows: {sheet.RowsCount}, Columns: {sheet.ColumnsCount}, Cells: {sheet.RowsCount * sheet.ColumnsCount})");

            AskKeyToExit();
        }

        private static void AskKeyToExit()
        {
            Console.Write("Press Any Key to exit...");
            Console.ReadKey();
        }
    }
}
