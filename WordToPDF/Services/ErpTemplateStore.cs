using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Options;
using WordToPDF.Models;

namespace WordToPDF.Services
{
    public sealed record StoredTemplate(int Id, string TemplateName);

    public sealed class ErpTemplateStore
    {
        private readonly DigitalInvoiceOptions _options;
        private readonly IConfiguration _configuration;

        public ErpTemplateStore(IOptions<DigitalInvoiceOptions> options, IConfiguration configuration)
        {
            _options = options.Value;
            _configuration = configuration;
        }

        public async Task<StoredTemplate> SaveAsync(string templateName, byte[] fileBytes, bool overwrite)
        {
            await using var connection = CreateConnection();
            await EnsureTableAsync(connection);

            var existingId = await GetIdByNameAsync(connection, templateName);
            if (existingId.HasValue)
            {
                if (!overwrite)
                {
                    throw new InvalidOperationException(
                        $"Template '{templateName}' already exists. Set overwrite=true to replace it.");
                }

                await using var update = new SqlCommand(
                    "UPDATE WordTemplates SET FileBytes = @bytes WHERE Id = @id", connection);
                update.Parameters.Add("@bytes", SqlDbType.VarBinary, -1).Value = fileBytes;
                update.Parameters.Add("@id", SqlDbType.Int).Value = existingId.Value;
                await update.ExecuteNonQueryAsync();

                return new StoredTemplate(existingId.Value, templateName);
            }

            await using var insert = new SqlCommand(
                "INSERT INTO WordTemplates (TemplateName, FileBytes) OUTPUT INSERTED.Id VALUES (@name, @bytes)",
                connection);
            insert.Parameters.Add("@name", SqlDbType.NVarChar, 255).Value = templateName;
            insert.Parameters.Add("@bytes", SqlDbType.VarBinary, -1).Value = fileBytes;

            var id = (int)(await insert.ExecuteScalarAsync())!;
            return new StoredTemplate(id, templateName);
        }

        public async Task<byte[]?> GetBytesAsync(int templateId)
        {
            await using var connection = CreateConnection();
            await EnsureTableAsync(connection);

            await using var command = new SqlCommand(
                "SELECT FileBytes FROM WordTemplates WHERE Id = @id", connection);
            command.Parameters.Add("@id", SqlDbType.Int).Value = templateId;

            var result = await command.ExecuteScalarAsync();
            return result == null ? null : (byte[])result;
        }

        private SqlConnection CreateConnection()
        {
            var connectionString = _configuration.GetConnectionString(_options.SqlServerConnectionStringName)
                ?? throw new InvalidOperationException(
                    $"Connection string '{_options.SqlServerConnectionStringName}' was not found.");

            return new SqlConnection(connectionString);
        }

        private static async Task EnsureTableAsync(SqlConnection connection)
        {
            const string createSql = @"
IF OBJECT_ID(N'dbo.WordTemplates', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.WordTemplates (
        Id           INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        TemplateName NVARCHAR(255)     NOT NULL,
        FileBytes    VARBINARY(MAX)    NOT NULL,
        CreatedAt    DATETIME2         NOT NULL CONSTRAINT DF_WordTemplates_CreatedAt DEFAULT SYSUTCDATETIME()
    );
END";
            await using var command = new SqlCommand(createSql, connection);
            await connection.OpenAsync();
            await command.ExecuteNonQueryAsync();
        }

        private static async Task<int?> GetIdByNameAsync(SqlConnection connection, string templateName)
        {
            await using var command = new SqlCommand(
                "SELECT Id FROM WordTemplates WHERE TemplateName = @name", connection);
            command.Parameters.Add("@name", SqlDbType.NVarChar, 255).Value = templateName;

            var result = await command.ExecuteScalarAsync();
            return result == null ? null : (int)result;
        }
    }
}
