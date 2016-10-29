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
    /// <summary>
    /// Base nonspecific Excel file reader. Reads pretty much any type of Excel file.
    /// </summary>
    public class ExcelFile
    {
        private readonly DataTable dt;

        public DataTable DT
        {
            get { return dt; }            
        }

                        
        /// <summary>
        /// Returns the contents of a specific cell in the Excel file.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <returns>Cell contents as String</returns>
        public string GetCellContents(int column, int row)
        {
            return DT.Rows[row].Field<string>(column);
        } 

        /// <summary>
        /// Exposes the number of rows in the Excel table
        /// </summary>
        public int RowCount
        {
            get { return DT.Rows.Count; }
            
        }

        /// <summary>
        /// Exposes the number of columns in the Excel table
        /// </summary>
        public int ColumnCount
        {
            get { return DT.Columns.Count; }
        
        }        

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="filePath"></param>
        public ExcelFile(string filePath)
        {
            
            var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader;

            //1. Reading Excel file
            if (Path.GetExtension(filePath).ToUpper() == ".XLS")
            {
                //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            //2. DataSet - The result of each spreadsheet will be created in the result.Tables
            var result = excelReader.AsDataSet();

            //3. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = false;

            dt = result.Tables[0];
            excelReader.Close();
        }        

    }

    /// <summary>
    /// An Excel file that contains schedule information. It is assumed that at least one of the columns contains a job number (does not need to be unique so that a job can be scheduled more than once) and at least one column contains a DateTime.
    /// </summary>
    
}
