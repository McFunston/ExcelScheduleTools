using ExcelDataGrabber;
using ExcelScheduleTools;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelScheduleToolsTests
{
    public static class Mock
    {
        public static ExcelFile returnFakeExcel()
        {
            var testTable = new DataTable();

            testTable.Columns.Clear();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Columns.Add();

            testTable.Rows.Add();
            testTable.Rows.Add();
            testTable.Rows.Add();

            testTable.Rows[0][0] = "111111";
            testTable.Rows[0][1] = "unuseable data";
            testTable.Rows[0][2] = "16/10/20 07:00";
            testTable.Rows[0][3] = "Company1";
            testTable.Rows[0][4] = "Task1";
            testTable.Rows[0][5] = 1.5;
            testTable.Rows[0][6] = 1000;
            testTable.Rows[0][7] = "Approved to Print";
            testTable.Rows[1][0] = "1111112";
            testTable.Rows[1][1] = "unuseable data2";
            testTable.Rows[1][2] = "17/11/21 08:00";
            testTable.Rows[1][3] = "Company2";
            testTable.Rows[1][4] = "Task2";
            testTable.Rows[1][5] = 2.5;
            testTable.Rows[1][6] = 2000;
            testTable.Rows[1][7] = "";
            testTable.Rows[2][0] = "1111112";
            testTable.Rows[2][1] = "unuseable data3";
            testTable.Rows[2][2] = "18/12/22 09:00";
            testTable.Rows[2][3] = "Company3";
            testTable.Rows[2][4] = "Task3";
            testTable.Rows[2][5] = 3.5;
            testTable.Rows[2][6] = 3000;
            testTable.Rows[1][7] = "Yes";

            var mockExcelFile = new ExcelFile(testTable);
            
            return mockExcelFile;
            }
        public static ExcelSchedule returnFakeSchedule()
        {
            var testTable = new DataTable();

            testTable.Columns.Clear();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Columns.Add();
            testTable.Rows.Add();
            testTable.Rows.Add();
            testTable.Rows.Add();

            testTable.Rows[0][0] = "111111";
            testTable.Rows[0][1] = "unuseable data";
            testTable.Rows[0][2] = "16/10/20 07:00";
            testTable.Rows[1][0] = "111112";
            testTable.Rows[1][1] = "unuseable data2";
            testTable.Rows[1][2] = "17/11/21 08:00";
            testTable.Rows[2][0] = "111112";
            testTable.Rows[2][1] = "unuseable data3";
            testTable.Rows[2][2] = "18/12/22 09:00";

            var fakeSchedule = new ExcelSchedule(testTable);
            fakeSchedule.JobNumberColumn = 0;
            return fakeSchedule;
        }


    
    }
}
