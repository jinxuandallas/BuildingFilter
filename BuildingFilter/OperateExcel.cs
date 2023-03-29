using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace BuildingFilter
{
    internal class OperateExcel
    {
        private string excelPath;
        private Excel.Application? myApp;
        private object missing = System.Reflection.Missing.Value;
        public OperateExcel()
        {
            excelPath = System.Environment.CurrentDirectory + @"\company.xlsx";
        }


        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hwnd, out int processid);
        public void OperateContent(List<List<string>> content)
        {
            myApp = new Excel.Application();
            Excel.Workbook workbook = myApp.Workbooks.Open(excelPath);
            Excel.Worksheet worksheet = (Excel.Worksheet)myApp.Worksheets.Add();


            Excel.Range range = (Excel.Range)worksheet.Cells[1, 5];
            range.HorizontalAlignment = Microsoft.Office.Core.XlHAlign.xlHAlignCenter;
            range.Value = DateTime.Now.Date.ToString("yyyy年MM月dd日");
            range.Font.Size = 30;

            int i = 2;
            foreach (var item in content)
            {
                item.Insert(0, (i - 1).ToString());
                i++;
                string[] i_arr = item.ToArray();
                Excel.Range row = worksheet.Range[worksheet.Cells[i, 2], worksheet.Cells[i, 2 + i_arr.Length - 1]];
                row.Value2 = i_arr;
            }
            //StreamReader streamReader = new StreamReader(content);

            //string? line = streamReader.ReadLine();
            //while (line != null)
            //{
            //    if (line.Contains("："))
            //    {

            //    }
            //    line = streamReader.ReadLine();

            //}

            //streamReader.Close();


            //画边框
            range = worksheet.Range[worksheet.Cells[3, 2], worksheet.Cells[worksheet.UsedRange.Rows.Count, 11]];
            range.Borders.Weight = 2;
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            //for(int j=0;j<14;j++)
            //{
            //    range= (Excel.Range)worksheet.Cells[20+j*2, 5];

            //    range.Borders.Weight = 2;
            //    range.Borders.LineStyle = j;
            //}


            //设置列宽
            int[] width = new int[] { 3, 40, 50, 20, 20, 30, 15, 12, 30, 15 };
            for (int j = 2; j < 12; j++)
            {
                range = (Excel.Range)worksheet.Cells[5, j];
                range.ColumnWidth = (object)width[j - 2];
            }

            workbook.Save();
            workbook.Close();

            myApp.Quit();

            int pid;

            GetWindowThreadProcessId(new IntPtr(myApp.Hwnd), out pid);

            System.Diagnostics.Process.GetProcessById(pid).Kill();


            //System.Runtime.InteropServices.Marshal.ReleaseComObject(myApp);
            //Process[] ExcelProcess = Process.GetProcessesByName("Excel");
            ////关闭进程
            //foreach (Process p in ExcelProcess)
            //{
            //    p.Kill();
            //}
            //myApp = null;
        }



        public void OperateContentAccumulation(List<List<string>> content)
        {
            myApp = new Excel.Application();
            Excel.Workbook workbook = myApp.Workbooks.Open(excelPath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];


            //Excel.Range range = (Excel.Range)worksheet.Cells[1, 5];
            //range.HorizontalAlignment = Microsoft.Office.Core.XlHAlign.xlHAlignCenter;
            //range.Value = DateTime.Now.Date.ToString("yyyy年MM月dd日");
            //range.Font.Size = 30;

            int usedrows = worksheet.UsedRange.Rows.Count;
            Excel.Range range = (Excel.Range)worksheet.Cells[usedrows, 2];
            string? num = range.Text.ToString();
            int i;
            if (!int.TryParse(num, out i))
                i = 0;
            foreach (var item in content)
            {
                //添加序号
                i++;
                item.Insert(0, i.ToString());

                string[] i_arr = item.ToArray();
                Excel.Range row = worksheet.Range[worksheet.Cells[worksheet.UsedRange.Rows.Count+1, 2], worksheet.Cells[worksheet.UsedRange.Rows.Count + 1, 2 + i_arr.Length - 1]];
                row.Value2 = i_arr;
                
            }
            //StreamReader streamReader = new StreamReader(content);

            //string? line = streamReader.ReadLine();
            //while (line != null)
            //{
            //    if (line.Contains("："))
            //    {

            //    }
            //    line = streamReader.ReadLine();

            //}

            //streamReader.Close();


            //画边框
            range = worksheet.Range[worksheet.Cells[usedrows, 2], worksheet.Cells[worksheet.UsedRange.Rows.Count, 11]];
            range.Borders.Weight = 2;
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            //for(int j=0;j<14;j++)
            //{
            //    range= (Excel.Range)worksheet.Cells[20+j*2, 5];

            //    range.Borders.Weight = 2;
            //    range.Borders.LineStyle = j;
            //}


            //设置列宽
            int[] width = new int[] { 5, 40, 50, 20, 20, 30, 12, 30, 15, 15 };
            for (int j = 2; j < 12; j++)
            {
                range = (Excel.Range)worksheet.Cells[5, j];
                //range.EntireColumn.AutoFit();
                range.ColumnWidth = (object)width[j - 2];
            }

            workbook.Save();
            workbook.Close();

            myApp.Quit();

            int pid;

            GetWindowThreadProcessId(new IntPtr(myApp.Hwnd), out pid);

            System.Diagnostics.Process.GetProcessById(pid).Kill();



        }
        public void test()
        {
            myApp = new Excel.Application();
            Excel.Workbook workbook = myApp.Workbooks.Open(excelPath);
            Excel.Worksheet worksheet = (Excel.Worksheet)myApp.Worksheets.Add();


            Excel.Range range = (Excel.Range)worksheet.Cells[1, 5];
            range.HorizontalAlignment = Microsoft.Office.Core.XlHAlign.xlHAlignCenter;
            range.Value = DateTime.Now.Date.ToString("yyyy年MM月dd日");
            range.Font.Size = 30;
            workbook.Save();
            workbook.Close();

            myApp.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(myApp);
            int pid;

            GetWindowThreadProcessId(new IntPtr(myApp.Hwnd), out pid);

            System.Diagnostics.Process.GetProcessById(pid).Kill();
            //PublicMethod.Kill(myApp);

            //Process[] ExcelProcess = Process.GetProcessesByName("Excel");
            ////关闭进程
            //foreach (Process p in ExcelProcess)
            //{
            //    p.Kill();
            //}
            myApp = null;

        }


    }




}
