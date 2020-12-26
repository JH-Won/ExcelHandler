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
        private string originFilePath;
        private Excel.Application app;
        private Excel.Workbook book;
        private Excel.Worksheet sheet;
        private Excel.Range range;

        /// <summary>
        /// 엑셀 파일과 시트 번호로 ExcelSheet 객체를 생성한다
        /// 파일 경로 또는 시트 번호가 유효하지 않은 경우, 예외를 throw한다
        /// </summary>
        /// <param name="excelFile">full file path</param>
        /// <param name="sheetNum">sheet number to apply</param>
        public ExcelSheet(string excelFile, int sheetNum = 1)
        {
            ValidatePath(excelFile);

            originFilePath = excelFile;
            app = new Excel.Application();
            book = app.Workbooks.Open(excelFile);

            if (sheetNum < 1 || sheetNum > book.Sheets.Count)
                throw new InvalidConstructionException("Invalid Sheet Number");

            sheet = book.Sheets[sheetNum];
            range = sheet.UsedRange;
        }

        // COM 오브젝트이므로 직접 리소스 해제한다
        ~ExcelSheet()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            
            if (range != null)
                Marshal.ReleaseComObject(range);
            if (sheet != null)
                Marshal.ReleaseComObject(sheet);

            if (book != null)
            {
                book.Close();
                Marshal.ReleaseComObject(book);
            }
            if (app != null)
            {
                app.Quit();
                Marshal.ReleaseComObject(app);
            }
        }

        /// <summary>
        /// 행 개수
        /// </summary>
        public int RowsCount { get => range.Rows.Count; }
        /// <summary>
        /// 열 개수
        /// </summary>
        public int ColumnsCount { get => range.Columns.Count; }
        /// <summary>
        /// 유효 데이터 셀 개수
        /// </summary>
        public int DataCellCount
        {
            get
            {
                int cnt = 0;
                for (int i = 1; i <= RowsCount; i++)
                    for (int j = 1; j <= ColumnsCount; j++)
                        if (!(this[i, j] == null) && !string.IsNullOrEmpty(this[i, j].ToString()))
                            cnt++;
                return cnt;
            }
        }

        /// <summary>
        /// 셀 데이터를 읽거나 쓴다
        /// 인덱스는 1부터 시작하고, 범위 초과시 예외를 throw한다
        /// </summary>
        /// <param name="row">row index, 1부터 시작</param>
        /// <param name="column">column index, 1부터 시작</param>
        /// <returns></returns>
        public object this[int row, int column]
        {
            get
            {
                ValidateIndex(row, column);
                try
                {
                    return range.Cells[row, column].Value2;
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

        /// <summary>
        /// 해당 시트의 모든 데이터를 DataTable로 만들어 반환한다
        /// 데이터 타입은 모두 Object타입으로 반환한다
        /// </summary>
        /// <returns>반환된 데이터테이블</returns>
        public DataTable GetFullDataTable()
        {
            DataTable ret = new DataTable();

            // 컬럼별로 하나씩 개별 스레드로 작업
            // 작업 끝나면 Merge (async - await 활용)

            return ret;
        }

        /// <summary>
        /// 첫번째 열을 컬럼 이름으로 하는 컬럼 이름 리스트를 반환한다
        /// </summary>
        /// <param name="isHeaderColumn">true if header is column names</param>
        /// <returns></returns>
        public IList<string> GetColumnNames(bool isHeaderColumn = true)
        {
            List<string> ret = new List<string>();

            if (isHeaderColumn)
            {
                for (int i = 1; i <= ColumnsCount; i++)
                    ret.Add(this[1, i].ToString());
            }

            return ret;            
        }

        /// <summary>
        /// 현재 열어둔 ExcelSheet의 변경사항을 저장한다
        /// </summary>
        public void Save()
        {
            this.book.Save();
        }

        /// <summary>
        /// 현재 시트의 복사본을 새 파일 경로에 저장한다
        /// 경로가 유효하지 않은 경우, 예외를 throw하고 저장되지 않는다
        /// </summary>
        /// <param name="filePath">저장할 파일 경로</param>
        public void SaveAs(string filePath)
        {
            ValidatePath(filePath);

            this.book.SaveCopyAs(filePath);
        }

        #region private methods
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
        #endregion
    }
}
