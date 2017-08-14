using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CFTCWorkSheet
{
    using Microsoft.Office.Interop.Excel;
    using System.Runtime.InteropServices;

    public class ExcelOperator
    {
        public ExcelOperator()
        {
            FilePath = "CFTC.xlsx";
            InitExcelApplication();
        }
        public string FilePath { get; }
        private Application excelApp;
        private Workbook xlsWorkbook;
        private readonly string DataSheet = "CFTC";
        private readonly string AnaySheet = "Analysis";
        private readonly int firstRow = 6;
        private int firstCol = 12;
        private Worksheet xlsWorkSheet;
        private Worksheet anayWorkSheet;
        
        private int _rowsCount;
        private int _colsCount;
        private void InitExcelApplication()
        {
            //excelApp = new Application();
            //excelApp.Visible = false;
            //excelApp.DisplayAlerts = false;
            //xlsWorkbook = excelApp.Workbooks.Open(FilePath, System.Type.Missing, System.Type.Missing, System.Type.Missing,
            //        System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
            //        System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
            //        System.Type.Missing);
            excelApp = Globals.CFTC.Application;
            xlsWorkbook = excelApp.ActiveWorkbook;
            xlsWorkSheet = xlsWorkbook.Worksheets[DataSheet];
           
           // xlsWorkSheet = (Worksheet)xlsWorkbook.Worksheets[DataSheet];
            _rowsCount = xlsWorkSheet.UsedRange.Rows.Count;
            _colsCount = xlsWorkSheet.UsedRange.Columns.Count;
        }

        public bool IsNeedUpdate(DateTime date)
        {
            //if today is saterday or sunday, get this week's tuesday date, or return last Tuesday date.
            //DateTime dt = getWeekUpOfDate(DateTime.Today, DayOfWeek.Tuesday, -1);
            DateTime thisDate = GetWeekUpOfDate();
            if (!date.Equals(thisDate))
            {
                return false;
            }
            DateTime temp = ((Range)xlsWorkSheet.Cells[6, 1]).Value;

            return temp.Equals(thisDate.AddDays(-7));
           
        }

        public DateTime GetWeekUpOfDate()
        {
            DateTime dtToday = DateTime.Now.Date;
            int wd1 = (int)DayOfWeek.Tuesday;
            int wd2 = (int)dtToday.DayOfWeek;

            if (dtToday.DayOfWeek == DayOfWeek.Saturday)
            {
                return  dtToday.AddDays(wd1 - wd2);
            }
            else
            {
                return dtToday.AddDays(7 * -1 - wd2 + wd1);
            }
        }

        public bool UpdateData(ref List<int> lst)
        {
            try
            {
                var rang = (Range)xlsWorkSheet.Rows[firstRow, Type.Missing];
                rang.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                foreach (int data in lst)
                {
                    xlsWorkSheet.Cells[firstRow, firstCol++] = data;
                }

                for (int index = 1; index < 10; index++)
                {
                    string formula = xlsWorkSheet.Rows[firstRow + 1].Columns[index].Formula;
                    xlsWorkSheet.Cells[index][firstRow] = excelApp.ConvertFormula(formula,
                        XlReferenceStyle.xlA1, XlReferenceStyle.xlR1C1,
                        XlReferenceType.xlRelative, xlsWorkSheet.Cells[index][firstRow + 1]);
                }
                AddAnaySheet();
                xlsWorkbook.Save();
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
               // CloseExcel(excelApp, xlsWorkbook);
            }

        }

        private void AddAnaySheet()
        {
            anayWorkSheet = Globals.ThisWorkbook.Worksheets[AnaySheet];

            var rang = (Range)anayWorkSheet.Rows[firstRow, Type.Missing];
            rang.Insert(XlInsertShiftDirection.xlShiftDown, XlInsertFormatOrigin.xlFormatFromRightOrBelow);
            for (int index = 2; index < 11; index++)
            {
                string formula = anayWorkSheet.Rows[firstRow + 1].Columns[index].Formula;
                anayWorkSheet.Cells[index][firstRow] = excelApp.ConvertFormula(formula,
                    XlReferenceStyle.xlA1, XlReferenceStyle.xlR1C1,
                    XlReferenceType.xlRelative, anayWorkSheet.Cells[index][firstRow + 1]);
            }
        }

        ///// <summary>
        ///// 关闭Excel进程
        ///// </summary>
        //public class KeyMyExcelProcess
        //{
        //    [DllImport("User32.dll", CharSet = CharSet.Auto)]
        //    public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        //    public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        //    {
        //        try
        //        {
        //            IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口
        //            int k = 0;
        //            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
        //            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
        //            p.Kill();     //关闭进程k
        //        }
        //        catch (System.Exception ex)
        //        {
        //            throw ex;
        //        }
        //    }
        //}


        ////关闭打开的Excel方法
        //public void CloseExcel(Application ExcelApplication, Workbook ExcelWorkbook)
        //{
        //    ExcelWorkbook.Close(false, Type.Missing, Type.Missing);
        //    ExcelWorkbook = null;
        //    ExcelApplication.Quit();
        //    GC.Collect();
        //    KeyMyExcelProcess.Kill(ExcelApplication);
        //}
    }
}
