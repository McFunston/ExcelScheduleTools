﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataGrabber
{
    public class ExcelSchedule : ExcelFile
    {
        public ExcelSchedule(string filePath) : base(filePath)
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
    }
}