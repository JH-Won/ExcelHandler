using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelExceptions;
using System.Data;

namespace ExcelHandler
{
    public sealed class ExcelHandler
    {
        public DataTable FullDataTable { get; private set; }
        //public int RowCount { get => }
        //public int ColumnCount { get => }

        private Excel.Application app;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file">파일 경로</param>
        public ExcelHandler(string file)
        {
            if (string.IsNullOrEmpty(file) || !File.Exists(file))
                throw new InvalidConstructionException();

            SetFullDataTable();
            SetColumns();
        }

        private void SetColumns()
        {
            throw new NotImplementedException();
        }

        private void SetFullDataTable()
        {
            throw new NotImplementedException();
        }
    }
}
