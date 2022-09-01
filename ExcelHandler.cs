using Sylvan.Data;
using Sylvan.Data.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;

namespace ReadExcelUsingSylvan
{

    internal static class ExcelHandler
    {
        public static DataTable SylvanReadExcelFile(string filePath)
        {
            var theData = new DataTable();

            var schema = new CustomSchema()
                .Add<int>("RowId")
                .Add<string>("First Name")
                .Add<string>("Last Name")
                .Add<string>("Gender")
                .Add<string>("Country")
                .Add<int>("Age")
                .Add<string>("DateString")
                .Add<int>("Id");

            using var edr = ExcelDataReader.Create(filePath, new ExcelDataReaderOptions
            {
                GetErrorAsNull = true,
                Schema = schema,
            });

            var idx = edr.GetOrdinal("DateString");

            // transform the raw ExcelDataReader to fix/workaround a couple issues
            var dr =
                edr
                // there is a blank row in the file. Skip it
                .Where(r => !r.IsDBNull(0))
                // attach a new column that
                .WithColumns(new CustomDataColumn<DateTime>("Date", r => DateTime.ParseExact(r.GetString(idx), "dd/MM/yyyy", null)));

            // finally, select everything excep the DateString column
            dr = dr.Select(d => d.GetColumnSchema().Where(c => c.ColumnName != "DateString").Select(c => c.ColumnOrdinal!.Value).ToArray());


            theData.Load(dr, LoadOption.PreserveChanges);

            return theData;
        }

        sealed class CustomSchema : IExcelSchemaProvider
        {
            List<Column> columns;

            public CustomSchema()
            {
                this.columns = new List<Column>();
            }

            public CustomSchema Add<T>(string name)
            {
                this.columns.Add(new Column(name, columns.Count, typeof(T)));
                return this;
            }

            class Column : DbColumn
            {
                public Column(string name, int ordinal, Type type)
                {
                    this.ColumnName = name;
                    this.ColumnOrdinal = ordinal;
                    this.DataType = type;
                }
            }

            public int GetFieldCount(ExcelDataReader reader)
            {
                // this handles the empty header row, which would otherwise be detected as 0 columns.
                return this.columns.Count;
            }

            public DbColumn? GetColumn(string sheetName, string? name, int ordinal)
            {
                // match by oridnal only, ignore any header in the file, which would include the "bad" headers.
                return ordinal < columns.Count ? columns[ordinal] : null;
            }

            public bool HasHeaders(string sheetName)
            {
                // the first row in your file is always a header (non-data) row.
                return true;
            }
        }

    }
}