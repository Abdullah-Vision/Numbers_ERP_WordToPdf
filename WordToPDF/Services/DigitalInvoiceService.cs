using Microsoft.Extensions.Options;
using WordToPDF.Models;

namespace WordToPDF.Services
{
    public interface IDigitalInvoiceService
    {
        Task<InvoiceDataSet> GetInvoiceDataAsync(int invoiceId, string dbType);
    }

    public sealed class DigitalInvoiceService : IDigitalInvoiceService
    {
        private readonly DigitalInvoiceOptions _options;
        private readonly IConfiguration _configuration;

        public DigitalInvoiceService(IOptions<DigitalInvoiceOptions> options, IConfiguration configuration)
        {
            _options = options.Value;
            _configuration = configuration;
        }

        public Task<InvoiceDataSet> GetInvoiceDataAsync(int invoiceId, string dbType)
        {
            var normalized = NormalizeDbType(dbType) ?? throw new InvalidOperationException(
                $"Unsupported dbType '{dbType}'. Use 'sqlserver' or 'oracle'.");

            return normalized == "oracle"
                ? CreateOracleExecutor().ExecuteAsync(invoiceId)
                : CreateSqlServerExecutor().ExecuteAsync(invoiceId);
        }

        public static string? NormalizeDbType(string? dbType)
        {
            return (dbType ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                "" or "sqlserver" or "sql" or "mssql" => "sqlserver",
                "oracle" => "oracle",
                _ => null
            };
        }

        private SqlServerInvoiceExecutor CreateSqlServerExecutor()
        {
            var connectionString = _configuration.GetConnectionString(_options.SqlServerConnectionStringName)
                ?? throw new InvalidOperationException(
                    $"Connection string '{_options.SqlServerConnectionStringName}' was not found.");

            return new SqlServerInvoiceExecutor(connectionString, _options.SqlServerProcedure);
        }

        private OracleInvoiceExecutor CreateOracleExecutor()
        {
            var connectionString = _configuration.GetConnectionString(_options.OracleConnectionStringName)
                ?? throw new InvalidOperationException(
                    $"Connection string '{_options.OracleConnectionStringName}' was not found.");

            return new OracleInvoiceExecutor(connectionString, _options.OracleProcedure);
        }
    }
}
