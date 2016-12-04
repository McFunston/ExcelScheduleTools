using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScheduleTools
{
    class Program
    {
        static void Main(string[] args)
        {
            var xf = new ExcelSchedule(DataGrabber.GrabExcelFile("c:\\Huntclub.xls"));

            xf.JobNumberColumn = 1;
            

            foreach (var JN in xf.ReturnUniqueJobNumbers())
            {
                Console.WriteLine(JN);
            }
            Console.ReadKey();


        }
    }
}
