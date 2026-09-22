using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Hosting;
using WordToPDF.Models;

namespace WordToPDF.Services
{
    public sealed record StoredTemplate(int Id, string TemplateName, string FilePath);

    public sealed class ErpTemplateStore
    {
        private readonly DigitalInvoiceOptions _options;
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _environment;

        public ErpTemplateStore(
            IOptions<DigitalInvoiceOptions> options,
            IConfiguration configuration,
            IWebHostEnvironment environment)
        {
            _options = options.Value;
            _configuration = configuration;
            _environment = environment;
        }

        public async Task<StoredTemplate> SaveAsync(string templateName, byte[] fileBytes, bool overwrite)
        {
            var filePath = WriteTemplateFile(templateName, fileBytes);

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
                    "UPDATE WordTemplates SET FileBytes = @bytes, FilePath = @path WHERE Id = @id", connection);
                update.Parameters.Add("@bytes", SqlDbType.VarBinary, -1).Value = fileBytes;
                update.Parameters.Add("@path", SqlDbType.NVarChar, 500).Value = filePath;
                update.Parameters.Add("@id", SqlDbType.Int).Value = existingId.Value;
                await update.ExecuteNonQueryAsync();

                return new StoredTemplate(existingId.Value, templateName, filePath);
            }

            await using var insert = new SqlCommand(
                "INSERT INTO WordTemplates (TemplateName, FileBytes, FilePath) OUTPUT INSERTED.Id VALUES (@name, @bytes, @path)",
                connection);
            insert.Parameters.Add("@name", SqlDbType.NVarChar, 255).Value = templateName;
            insert.Parameters.Add("@bytes", SqlDbType.VarBinary, -1).Value = fileBytes;
            insert.Parameters.Add("@path", SqlDbType.NVarChar, 500).Value = filePath;

            var id = (int)(await insert.ExecuteScalarAsync())!;
            return new StoredTemplate(id, templateName, filePath);
        }

        public async Task<byte[]?> GetBytesAsync(int templateId)
        {
            await using var connection = CreateConnection();
            await EnsureTableAsync(connection);

            await using var command = new SqlCommand(
                "SELECT FileBytes, FilePath FROM WordTemplates WHERE Id = @id", connection);
            command.Parameters.Add("@id", SqlDbType.Int).Value = templateId;

            await using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync())
            {
                return null;
            }

            var filePath = reader.IsDBNull(1) ? null : reader.GetString(1);
            if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
            {
                return await File.ReadAllBytesAsync(filePath);
            }

            return reader.IsDBNull(0) ? null : (byte[])reader["FileBytes"];
        }

        private string WriteTemplateFile(string templateName, byte[] fileBytes)
        {
            var storagePath = Path.IsPathRooted(_options.TemplateStoragePath)
                ? _options.TemplateStoragePath
                : Path.Combine(_environment.ContentRootPath, _options.TemplateStoragePath);

            Directory.CreateDirectory(storagePath);

            var safeName = string.Join("_",
                templateName.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries));

            if (string.IsNullOrWhiteSpace(safeName))
            {
                safeName = "template";
            }

            var fullPath = Path.Combine(storagePath, $"{safeName}.docx");

            File.WriteAllBytes(fullPath, fileBytes);
            return fullPath;
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
            const string ensureSql = @"
IF OBJECT_ID(N'dbo.WordTemplates', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.WordTemplates (
        Id           INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        TemplateName NVARCHAR(255)     NOT NULL,
        FileBytes    VARBINARY(MAX)    NOT NULL,
        FilePath     NVARCHAR(500)     NULL,
        CreatedAt    DATETIME2         NOT NULL CONSTRAINT DF_WordTemplates_CreatedAt DEFAULT SYSUTCDATETIME()
    );
END
ELSE IF COL_LENGTH(N'dbo.WordTemplates', N'FilePath') IS NULL
BEGIN
    ALTER TABLE dbo.WordTemplates ADD FilePath NVARCHAR(500) NULL;
END";
            await using var command = new SqlCommand(ensureSql, connection);
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
