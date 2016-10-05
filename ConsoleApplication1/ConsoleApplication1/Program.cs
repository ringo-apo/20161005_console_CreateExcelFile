using Microsoft.Office.Interop.Excel;
using System;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            string a = Console.ReadLine();

            string ExcelBookFileName = @"D:\test";
            Application ExcelApp = new Application();
            ExcelApp.Visible = false;
            Workbook wb = ExcelApp.Workbooks.Add();
            Worksheet ws1 = wb.Sheets[1];
            ws1.Select(Type.Missing);
            ws1.Cells[1, 1] = a;
            wb.SaveAs(ExcelBookFileName);
            wb.Close(false);
            ExcelApp.Quit();

            Console.ReadKey(); //文字を押すと終了
        }
    }
}
