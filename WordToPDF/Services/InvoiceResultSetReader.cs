using System.Data.Common;
using WordToPDF.Models;

namespace WordToPDF.Services
{
    public static class InvoiceResultSetReader
    {
        public static async Task<InvoiceDataSet> ReadAsync(DbDataReader reader)
        {
            var result = new InvoiceDataSet();

            foreach (var column in reader.GetColumnSchema())
            {
                result.Columns.Add(column.ColumnName);
            }

            while (await reader.ReadAsync())
            {
                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var value = reader.GetValue(i);
                    row[reader.GetName(i)] = value is DBNull ? null : value;
                }

                result.Rows.Add(row);
            }

            return result;
        }
    }
}
