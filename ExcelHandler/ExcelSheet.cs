using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ExcelExceptions;
using System.Data;
using System.Runtime.InteropServices;

namespace ExcelHandler
{
    public sealed class ExcelSheet
    {
        public int RowsCount { get => range.Rows.Count; }
        public int ColumnsCount { get => range.Columns.Count; }

        public object this[int row, int column]
        {
            get
            {
                ValidateIndex(row, column);
                try
                {
                    return range.Cells[row, column];
                }
                catch(Exception e)
                {
                    throw e;
                }
            }
            set
            {
                ValidateIndex(row, column);
                try
                {
                    range.Cells[row, column] = value;
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        private string originFilePath;

        private Excel.Application app;
        private Excel.Workbook book;
        private Excel.Worksheet sheet;
        private Excel.Range range;


        public ExcelSheet(string excelFile, int sheetNum = 1)
        {
            originFilePath = excelFile;
        }

        ~ExcelSheet()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(sheet);

            book.Close();
            Marshal.ReleaseComObject(book);

            app.Quit();
            Marshal.ReleaseComObject(app);
        }

        public DataTable GetFullDataTable()
        {
            DataTable ret = new DataTable();

            // 컬럼별로 하나씩 개별 스레드로 작업
            // 작업 끝나면 Merge (async - await 활용)

            return ret;
        }

        public IList<string> GetColumnNames(bool isHeaderColumn = true)
        {
            if (isHeaderColumn)
            {

            }

            return new List<string>();
        }

        public void Save(string filePath)
        {
            ValidatePath(filePath);

            try
            {
                // 저장한다
            }
            catch(Exception e)
            {
                throw e;
            }
        }

        private void ValidatePath(string filePath)
        {
            if (filePath == originFilePath)
                throw new InvalidPathException();

            FileInfo fileInfo = null;
            try
            {
                fileInfo = new FileInfo(filePath);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                if (ReferenceEquals(fileInfo, null))
                    throw new InvalidPathException();
            }
        }

        private void ValidateIndex(int row, int column)
        {
            if (row < 1 || row > RowsCount)
                throw new IndexOutOfRangeException("row index out of range, excel sheet's cell index starts with 1(is not zero-based) and ends with row or coulmn count");
            else if (column < 1 || column > ColumnsCount)
                throw new IndexOutOfRangeException("column index out of range, excel sheet's cell index starts with 1(is not zero-based) and ends with row or coulmn count");
        }
    }
}
