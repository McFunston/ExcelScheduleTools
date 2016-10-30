using ExcelDataGrabber;
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
        testTable.Rows.Add();
        testTable.Rows.Add();
        testTable.Rows.Add();

        testTable.Rows[0][0] = "111111";
        testTable.Rows[0][1] = "unuseable data";
        testTable.Rows[0][2] = "16/10/20 07:00";
        testTable.Rows[1][0] = "1111112";
        testTable.Rows[1][1] = "unuseable data2";
        testTable.Rows[1][2] = "17/11/21 08:00";
        testTable.Rows[2][0] = "1111112";
        testTable.Rows[2][1] = "unuseable data3";
        testTable.Rows[2][2] = "18/12/22 09:00";

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
            testTable.Rows[1][0] = "1111112";
            testTable.Rows[1][1] = "unuseable data2";
            testTable.Rows[1][2] = "17/11/21 08:00";
            testTable.Rows[2][0] = "1111112";
            testTable.Rows[2][1] = "unuseable data3";
            testTable.Rows[2][2] = "18/12/22 09:00";

            var fakeSchedule = new ExcelSchedule(testTable);
            fakeSchedule.JobNumberColumn = 0;
            return fakeSchedule;
        }


    
    }
}
