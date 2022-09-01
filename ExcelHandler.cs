using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sylvan.Data.Excel;

namespace ReadExcelUsingSylvan
{
    internal static class ExcelHandler
    {
        public static DataTable FixHeaders(DataTable dt)
        {
            foreach (DataColumn column in dt.Columns)
            {
                //these are a few (of many) illegal characters that aren't accepted in the System.Windows.Controls.DataGrid Column Header 
                column.ColumnName = column.ColumnName.Replace(".", "_").Replace("#", "").Replace("/", " ").Replace(":", " ");
            }

            return dt;
        }

        public static DataTable SylvanReadExcelFile(string filePath)
        {
            var theData = new DataTable();

            using var edr = ExcelDataReader.Create(filePath, new ExcelDataReaderOptions
            {
                GetErrorAsNull = true,
            });

            theData.Load(edr, LoadOption.PreserveChanges);

            return theData;
        }
    }
}
