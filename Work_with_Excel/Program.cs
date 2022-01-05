using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace Work_with_Excel
{
    class Excel
    {
        string path = "";
        _Application excel = new Application();
        Workbook workbook;
        Worksheet worksheet;

        public Excel(string path, int sheet)
        {
            this.path = path;
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheet];
        }

        public Excel(string path)
        {
            int sheet = 1;
            this.path = path;
            workbook = excel.Workbooks.Open(path);
            worksheet = workbook.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            if (worksheet.Cells[i, j].Value2 != null)
                return worksheet.Cells[i, j].Value2.ToString();
            else return "Пусто :(";
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(DateTime.Now);
            //Excel excel = new Excel("C:\\Users\\USER\\source\\repos\\Master\\Test.xlsx", 1);

            //for (int i = 1; i < 10; i++)            
            //    Console.WriteLine(excel.ReadCell(1, i));
            
            Console.ReadKey();

        }
    }
}
