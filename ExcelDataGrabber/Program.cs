using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataGrabber
{
    class Program
    {
        static void Main(string[] args)
        {
            var xf = new ExcelFile("c:\\Huntclub.xls");
            xf.JobNumberColumn = 1;


            
            Console.ReadKey();


        }
    }
}
