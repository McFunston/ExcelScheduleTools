﻿using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataGrabber
{
    public class ExcelFile
    {
        public DataTable DT { get; set; }
        private int rowCount;

        public int RowCount
        {
            get { return DT.Rows.Count; }
            
        }
        private int columnCount;

        public int ColumnCount
        {
            get { return DT.Columns.Count; }
        
        }
        public int JobNumberColumn { get; set; }

        public ExcelFile(string filePath)
        {
            
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
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
            DataSet result = excelReader.AsDataSet();

            //3. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = false;

            DT = result.Tables[0];
            excelReader.Close();
        }
        
    }
}
