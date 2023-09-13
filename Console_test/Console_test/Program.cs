using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace Console_test
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application _app;
            Excel.Worksheet _sheet;
            Excel.Workbook _book;
            Excel.Range _range;


            string txt_path = System.AppDomain.CurrentDomain.BaseDirectory + "\\test.txt";
            ////Pass the filepath and filename to the StreamWriter Constructor
            //StreamWriter sw = new StreamWriter(txt_path);
            ////Write a line of text
            //sw.WriteLine("AAA\tBBB\tCCC");
            //sw.WriteLine("DDD\tEEE\tFFF");
            ////Close the file
            //sw.Close();


            string path = System.AppDomain.CurrentDomain.BaseDirectory + "\\test.xlsx";
            _app = new Excel.Application();
            _app.Visible = true;
            _book = _app.Workbooks.Open(path);
            _sheet = (Excel.Worksheet)_app.ActiveSheet;


            StreamReader sr = new StreamReader(txt_path);
            string line = sr.ReadToEnd();
            int row = 1;
            
            string[] str_ar = line.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            foreach(string tmp in str_ar)
            {
                _range = _sheet.Range["A" + row, "C" + (row + 1)];
                _range.Value = str_ar[row - 1].Split('\t');
                row += 1;
            }

            Console.ReadLine();
            //close the file
            sr.Close();
            _book.Save();
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();


        }
    }
}
