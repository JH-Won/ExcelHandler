using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Data;

namespace ExcelHandler
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Program Start");
            string file = Console.ReadLine();
                
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

            Console.Write($"Column Names(Header) : {sheet.RowsCount}");
            foreach (var col in sheet.GetColumnNames())
            {
                Console.Write($"{col}, ");
            }

            for (int i = 0; i < sheet.RowsCount; i++) 
            {
                for (int j = 0; j < sheet.ColumnsCount; j++)
                {
                    Console.Write(sheet[i, j] + "\t");
                }
                Console.WriteLine();
            }

            Console.WriteLine("Converting DataTable...");
            timer.Start();
            using(DataTable dt = sheet.GetFullDataTable())
            {               
            }
            timer.Stop();
            Console.WriteLine("...Done!");
            Console.WriteLine($"time elapsed to convert DataTable : {timer.ElapsedMilliseconds} ms");

            AskKeyToExit();
        }

        private static void AskKeyToExit()
        {
            Console.Write("Press Any Key to exit...");
            Console.ReadKey();
        }
    }
}
