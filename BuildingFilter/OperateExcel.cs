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
                i++;
                string[] i_arr=item.ToArray();
                Excel.Range row = worksheet.Range[worksheet.Cells[i, 2], worksheet.Cells[i,2+i_arr.Length-1]];
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

        public void test()
        {
            myApp = new Excel.Application();
            Excel.Workbook workbook = myApp.Workbooks.Open(excelPath);
            Excel.Worksheet worksheet = (Excel.Worksheet)myApp.Worksheets.Add();


            Excel.Range range = (Excel.Range)worksheet.Cells[1,5];
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
