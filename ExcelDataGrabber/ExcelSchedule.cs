using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using ExcelDataGrabber;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ScheduleEntities;

namespace ExcelScheduleTools
{
    public class ExcelSchedule : ExcelFile
    {
        public ExcelSchedule(DataTable dataTable) : base(dataTable)
        {
        }

        public int JobNumberColumn { get; set; }

        /// <summary>
        /// Returns a list of job numbers (invoice numbers, docket numbers ... etc). Requires JobNumberColumn to be set or it will return Null
        /// </summary>
        /// <returns>List of strings</returns>
        public List<string> ReturnJobNumbers()
        {
            var JobsNumberList = new List<string>();
            int JN;

            if (JobNumberColumn == -1)
            {
                return null;
            }
            else
            {
                for (int i = 0; i < this.RowCount; i++)
                {
                    if (int.TryParse(GetCellContents(JobNumberColumn, i), out JN))
                    {
                        JobsNumberList.Add(DT.Rows[i][JobNumberColumn].ToString());
                    }

                }
                return JobsNumberList;
            }
        }

        /// <summary>
        /// Returns a list of only the unique job numbers in a schedule. Requires JobNumberColumn to be set or it will return Null.
        /// </summary>
        /// <returns>List of strings</returns>
        public List<string> ReturnUniqueJobNumbers()
        {
            var UniqueJobNumbers = new List<string>();

            if (JobNumberColumn == -1)
            {
                return null;
            }
            else
            {
                foreach (var JN in ReturnJobNumbers())
                {
                    if (!(UniqueJobNumbers.Contains(JN)))
                    {
                        UniqueJobNumbers.Add(JN);
                    }
                }
                return UniqueJobNumbers;
            }
        }

        /// <summary>
        /// Get the row number for the first instance of a job number 
        /// </summary>
        /// <param name="jobNumber"></param>
        /// <returns></returns>
        public int GetJobNumberRow(string jobNumber)
        {
            int i = 0;
            foreach (string JN in ReturnJobNumbers())
            {
                
                if (JN.Equals(jobNumber))
                {
                    return i;                                      
                }
                i++;
            }
            return i;
        }

        /// <summary>
        /// Return a proper DateTime from a Cell using format as the Format string
        /// </summary>
        /// <param name="column"></param>
        /// <param name="jobNumber"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public DateTime ReturnDateTimeFromCell(int column, string jobNumber, string format)
        {
            CultureInfo ci = new CultureInfo(CultureInfo.CurrentCulture.LCID);
            ci.Calendar.TwoDigitYearMax = 2099;
            DateTime dt = DateTime.ParseExact(GetCellContents(column, GetJobNumberRow(jobNumber)), format, ci);
            return dt;
        }
        public Schedule ConvertToSchedule(ColumnWorkUnitMap map)
        {
            Schedule schedule = new Schedule();

            return schedule;
        }
    }
}
