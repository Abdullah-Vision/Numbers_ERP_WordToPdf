using System.Data;
using System.Data.Common;
using Microsoft.Data.SqlClient;
using Oracle.ManagedDataAccess.Client;
using WordToPDF.Models;

namespace WordToPDF.Services
{
    public interface IInvoiceDbExecutor
    {
        Task<InvoiceDataSet> ExecuteAsync(int invoiceId);
    }

    public sealed class SqlServerInvoiceExecutor : IInvoiceDbExecutor
    {
        private readonly string _connectionString;
        private readonly string _procedureName;

        public SqlServerInvoiceExecutor(string connectionString, string procedureName)
        {
            _connectionString = connectionString;
            _procedureName = procedureName;
        }

        public async Task<InvoiceDataSet> ExecuteAsync(int invoiceId)
        {
            await using var connection = new SqlConnection(_connectionString);
            await using var command = connection.CreateCommand();
            command.CommandText = _procedureName;
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.Add("@id", SqlDbType.Int).Value = invoiceId;

            await connection.OpenAsync();
            using var reader = await command.ExecuteReaderAsync();
            return await InvoiceResultSetReader.ReadAsync(reader);
        }
    }

    public sealed class OracleInvoiceExecutor : IInvoiceDbExecutor
    {
        private readonly string _connectionString;
        private readonly string _procedureName;

        public OracleInvoiceExecutor(string connectionString, string procedureName)
        {
            _connectionString = connectionString;
            _procedureName = procedureName;
        }

        public async Task<InvoiceDataSet> ExecuteAsync(int invoiceId)
        {
            await using var connection = new OracleConnection(_connectionString);
            await using var command = connection.CreateCommand();
            command.CommandText = _procedureName;
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.Add(new OracleParameter("p_id", OracleDbType.Int32, invoiceId, ParameterDirection.Input));
            command.Parameters.Add(new OracleParameter("p_result", OracleDbType.RefCursor, ParameterDirection.Output));

            await connection.OpenAsync();
            using var reader = await command.ExecuteReaderAsync();
            return await InvoiceResultSetReader.ReadAsync(reader);
        }
    }
}
