# Digital Invoice Template Engine Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fill a Word template whose `{Placeholders}` match the aliases of stored procedure `SpArDigitalInvoices` at runtime (zero code change when the proc changes) and return a downloadable PDF.

**Architecture:** ADO.NET executors (SQL Server now, Oracle ready) load proc rows as `Dictionary<alias, value>`; a metadata-driven OpenXML filler detects "detail" columns (values that vary across rows), clones detail table rows per invoice line, fills master placeholders from the first row; LibreOffice headless converts DOCX to PDF.

**Tech Stack:** ASP.NET Core 8 (net8.0), DocumentFormat.OpenXml 3.3.0, Microsoft.Data.SqlClient 6.0.2, Oracle.ManagedDataAccess.Core 23.9.1, xUnit, LibreOffice `soffice.exe`.

**Spec:** `docs/superpowers/specs/2026-09-16-digital-invoice-template-engine-design.md`

## Global Constraints

- **NO GIT COMMITS by the implementing agent.** The user reviews every change and commits personally. Each task ends at "STOP — present changes to user"; the user commits.
- Target framework `net8.0`; `Nullable` and `ImplicitUsings` enabled (match `WordToPDF.csproj`).
- No code comments in source files (repo style).
- New service files go in `WordToPDF/Services/` with namespace `WordToPDF.Services`; new models in `WordToPDF/Models/` with namespace `WordToPDF.Models` (existing services in `Controllers/` keep their location; do not move them).
- Number/date formatting uses `CultureInfo.InvariantCulture` for deterministic output.
- Culture for dates default `dd-MMM-yy`; decimals default `N2`.
- Dictionaries of result values use `StringComparer.OrdinalIgnoreCase`.
- Run all CLI commands from repo root `D:\Wordtopdf\WordToPDF`.
- `SqlServerConnection` in appsettings currently points at `EmployeeDB`; the ERP DB containing `SpArDigitalInvoices` may differ — connection string name is configurable (`DigitalInvoice:SqlServerConnectionStringName`); flag this to the user at the end.

---

### Task 1: Test project + token parsing + value formatting

**Files:**
- Create: `WordToPDF.Tests/WordToPDF.Tests.csproj`
- Create: `WordToPDF.Tests/InvoicePlaceholdersTests.cs`
- Create: `WordToPDF/Services/InvoicePlaceholders.cs`
- Modify: `WordToPDF.sln` (add test project)

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `record InvoiceToken(string Name, string? Format, string Raw)` in `WordToPDF.Services`
  - `static List<InvoiceToken> InvoicePlaceholders.ParseTokens(string? text)`
  - `static string InvoicePlaceholders.FormatValue(object? value, string? format)`

- [ ] **Step 1: Create the xUnit test project and add it to the solution**

`WordToPDF.Tests/WordToPDF.Tests.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <IsPackable>false</IsPackable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.9.0" />
    <PackageReference Include="xunit" Version="2.7.0" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.5.7">
      <PrivateAssets>all</PrivateAssets>
      <IncludeAssets>runtime; build; native; contentfiles; analyzers; buildtransitive</IncludeAssets>
    </PackageReference>
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\WordToPDF\WordToPDF.csproj" />
  </ItemGroup>

</Project>
```

Command:

```bash
dotnet sln WordToPDF.sln add WordToPDF.Tests/WordToPDF.Tests.csproj
```

- [ ] **Step 2: Write the failing tests**

`WordToPDF.Tests/InvoicePlaceholdersTests.cs`:

```csharp
using WordToPDF.Services;
using Xunit;

namespace WordToPDF.Tests
{
    public class InvoiceTokenParsingTests
    {
        [Fact]
        public void SimpleToken_ParsesNameWithoutFormat()
        {
            var tokens = InvoicePlaceholders.ParseTokens("Invoice No: {InvoiceNo}");

            var token = Assert.Single(tokens);
            Assert.Equal("InvoiceNo", token.Name);
            Assert.Null(token.Format);
            Assert.Equal("{InvoiceNo}", token.Raw);
        }

        [Fact]
        public void TokenWithFormat_ParsesBothParts()
        {
            var tokens = InvoicePlaceholders.ParseTokens("Date: {InvoiceDate:dd-MMM-yy}");

            var token = Assert.Single(tokens);
            Assert.Equal("InvoiceDate", token.Name);
            Assert.Equal("dd-MMM-yy", token.Format);
        }

        [Fact]
        public void MultipleTokens_AreAllParsed()
        {
            var tokens = InvoicePlaceholders.ParseTokens("{InvoiceNo} due {InvoiceDate:yyyy-MM} end");

            Assert.Equal(2, tokens.Count);
            Assert.Equal("InvoiceNo", tokens[0].Name);
            Assert.Equal("InvoiceDate", tokens[1].Name);
            Assert.Equal("yyyy-MM", tokens[1].Format);
        }

        [Fact]
        public void TokenNameIsTrimmed()
        {
            var tokens = InvoicePlaceholders.ParseTokens("{ InvoiceNo }");

            Assert.Equal("InvoiceNo", Assert.Single(tokens).Name);
        }

        [Fact]
        public void TextWithoutTokens_ReturnsEmptyList()
        {
            Assert.Empty(InvoicePlaceholders.ParseTokens("no tokens here"));
            Assert.Empty(InvoicePlaceholders.ParseTokens(null));
        }
    }

    public class FormatValueTests
    {
        [Fact]
        public void NullValue_ReturnsEmptyString()
        {
            Assert.Equal(string.Empty, InvoicePlaceholders.FormatValue(null, null));
        }

        [Fact]
        public void DbNull_ReturnsEmptyString()
        {
            Assert.Equal(string.Empty, InvoicePlaceholders.FormatValue(DBNull.Value, null));
        }

        [Fact]
        public void ByteArray_ReturnsEmptyString()
        {
            Assert.Equal(string.Empty, InvoicePlaceholders.FormatValue(new byte[] { 1, 2 }, null));
        }

        [Fact]
        public void DateTime_UsesDefaultFormat()
        {
            var value = new DateTime(2026, 9, 14);
            Assert.Equal("14-Sep-26", InvoicePlaceholders.FormatValue(value, null));
        }

        [Fact]
        public void DateTime_UsesCustomFormat()
        {
            var value = new DateTime(2026, 9, 14);
            Assert.Equal("2026-09-14", InvoicePlaceholders.FormatValue(value, "yyyy-MM-dd"));
        }

        [Fact]
        public void Decimal_UsesN2Default()
        {
            Assert.Equal("1,234,567.89", InvoicePlaceholders.FormatValue(1234567.891m, null));
        }

        [Fact]
        public void Decimal_UsesCustomFormat()
        {
            Assert.Equal("1,271,187", InvoicePlaceholders.FormatValue(1271187m, "N0"));
        }

        [Fact]
        public void Integer_UsesN2Default()
        {
            Assert.Equal("18.00", InvoicePlaceholders.FormatValue(18, null));
        }

        [Fact]
        public void String_PassesThrough()
        {
            Assert.Equal("abc", InvoicePlaceholders.FormatValue("abc", null));
        }
    }
}
```

- [ ] **Step 3: Run tests to verify they fail**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: **build error** — `InvoicePlaceholders` does not exist.

- [ ] **Step 4: Implement minimal code**

`WordToPDF/Services/InvoicePlaceholders.cs`:

```csharp
using System.Globalization;
using System.Text.RegularExpressions;

namespace WordToPDF.Services
{
    public sealed record InvoiceToken(string Name, string? Format, string Raw);

    public static class InvoicePlaceholders
    {
        public static readonly Regex TokenPattern =
            new(@"\{(?<name>[^{}:]+)(?::(?<format>[^{}]+))?\}", RegexOptions.Compiled);

        public static List<InvoiceToken> ParseTokens(string? text)
        {
            var tokens = new List<InvoiceToken>();
            if (string.IsNullOrEmpty(text))
            {
                return tokens;
            }

            foreach (Match match in TokenPattern.Matches(text))
            {
                tokens.Add(new InvoiceToken(
                    match.Groups["name"].Value.Trim(),
                    match.Groups["format"].Success ? match.Groups["format"].Value.Trim() : null,
                    match.Value));
            }

            return tokens;
        }

        public static string FormatValue(object? value, string? format)
        {
            if (value is null || value is DBNull)
            {
                return string.Empty;
            }

            if (value is byte[])
            {
                return string.Empty;
            }

            if (value is DateTime dateTime)
            {
                return dateTime.ToString(
                    string.IsNullOrWhiteSpace(format) ? "dd-MMM-yy" : format,
                    CultureInfo.InvariantCulture);
            }

            if (value is IFormattable formattable)
            {
                return formattable.ToString(
                    string.IsNullOrWhiteSpace(format) ? "N2" : format,
                    CultureInfo.InvariantCulture);
            }

            return value.ToString() ?? string.Empty;
        }
    }
}
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: all tests PASS.

- [ ] **Step 6: STOP — present changes to the user for review. The user commits.**

---

### Task 2: Detail-column computation + detail-row detection

**Files:**
- Modify: `WordToPDF.Tests/InvoicePlaceholdersTests.cs` (add test classes)
- Modify: `WordToPDF/Services/InvoicePlaceholders.cs`
- Create: `WordToPDF/Models/InvoiceDataSet.cs`

**Interfaces:**
- Consumes: `InvoicePlaceholders.ParseTokens` (Task 1).
- Produces:
  - `class InvoiceDataSet { List<string> Columns; List<Dictionary<string, object?>> Rows; Dictionary<string, object?>? FirstRow; }` in `WordToPDF.Models`
  - `static HashSet<string> InvoicePlaceholders.GetDetailColumns(InvoiceDataSet data)`
  - `static bool InvoicePlaceholders.IsDetailRow(string rowText, HashSet<string> detailColumns)`

- [ ] **Step 1: Create the InvoiceDataSet model (no behavior, needed for tests)**

`WordToPDF/Models/InvoiceDataSet.cs`:

```csharp
namespace WordToPDF.Models
{
    public sealed class InvoiceDataSet
    {
        public List<string> Columns { get; set; } = new();

        public List<Dictionary<string, object?>> Rows { get; set; } = new();

        public Dictionary<string, object?>? FirstRow => Rows.Count > 0 ? Rows[0] : null;
    }
}
```

- [ ] **Step 2: Write the failing tests**

Append to `WordToPDF.Tests/InvoicePlaceholdersTests.cs` (inside the namespace, after `FormatValueTests`):

```csharp
    public class DetailColumnTests
    {
        private static InvoiceDataSet Data(params Dictionary<string, object?>[] rows)
        {
            var data = new InvoiceDataSet { Columns = { "Qty", "Rate", "CustomerName" } };
            data.Rows.AddRange(rows);
            return data;
        }

        [Fact]
        public void VaryingColumn_IsDetail()
        {
            var data = Data(
                new Dictionary<string, object?> { ["Qty"] = 1m, ["Rate"] = 100m, ["CustomerName"] = "Acme" },
                new Dictionary<string, object?> { ["Qty"] = 2m, ["Rate"] = 100m, ["CustomerName"] = "Acme" });

            var detail = InvoicePlaceholders.GetDetailColumns(data);

            Assert.Contains("Qty", detail);
        }

        [Fact]
        public void InvariantColumn_IsNotDetail()
        {
            var data = Data(
                new Dictionary<string, object?> { ["Qty"] = 1m, ["Rate"] = 100m, ["CustomerName"] = "Acme" },
                new Dictionary<string, object?> { ["Qty"] = 2m, ["Rate"] = 100m, ["CustomerName"] = "Acme" });

            var detail = InvoicePlaceholders.GetDetailColumns(data);

            Assert.DoesNotContain("Rate", detail);
            Assert.DoesNotContain("CustomerName", detail);
        }

        [Fact]
        public void NullVsValue_CountsAsVarying()
        {
            var data = Data(
                new Dictionary<string, object?> { ["Qty"] = 1m, ["Rate"] = 100m, ["CustomerName"] = "Acme" },
                new Dictionary<string, object?> { ["Qty"] = null, ["Rate"] = 100m, ["CustomerName"] = "Acme" });

            Assert.Contains("Qty", InvoicePlaceholders.GetDetailColumns(data));
        }

        [Fact]
        public void SingleRow_ProducesNoDetailColumns()
        {
            var data = Data(
                new Dictionary<string, object?> { ["Qty"] = 1m, ["Rate"] = 100m, ["CustomerName"] = "Acme" });

            Assert.Empty(InvoicePlaceholders.GetDetailColumns(data));
        }

        [Fact]
        public void DetailRow_ContainingDetailColumnToken_ReturnsTrue()
        {
            var detail = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Qty" };

            Assert.True(InvoicePlaceholders.IsDetailRow("{Qty} | {CustomerName}", detail));
        }

        [Fact]
        public void MasterRow_WithOnlyMasterTokens_ReturnsFalse()
        {
            var detail = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Qty" };

            Assert.False(InvoicePlaceholders.IsDetailRow("Total: {GrandTotal}", detail));
        }

        [Fact]
        public void DetailRow_MatchingIsCaseInsensitive()
        {
            var detail = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Qty" };

            Assert.True(InvoicePlaceholders.IsDetailRow("{qty}", detail));
        }
    }
```

- [ ] **Step 3: Run tests to verify they fail**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: build error — `GetDetailColumns` / `IsDetailRow` do not exist.

- [ ] **Step 4: Implement minimal code**

Append inside `InvoicePlaceholders` in `WordToPDF/Services/InvoicePlaceholders.cs`:

```csharp
        public static HashSet<string> GetDetailColumns(InvoiceDataSet data)
        {
            var detailColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (data.Rows.Count < 2)
            {
                return detailColumns;
            }

            foreach (var column in data.Columns)
            {
                string? first = null;
                var varies = false;
                foreach (var row in data.Rows)
                {
                    var key = ValueKey(row.TryGetValue(column, out var value) ? value : null);
                    if (first is null)
                    {
                        first = key;
                    }
                    else if (!string.Equals(first, key, StringComparison.Ordinal))
                    {
                        varies = true;
                        break;
                    }
                }

                if (varies)
                {
                    detailColumns.Add(column);
                }
            }

            return detailColumns;
        }

        public static bool IsDetailRow(string rowText, HashSet<string> detailColumns)
        {
            if (detailColumns.Count == 0 || string.IsNullOrEmpty(rowText))
            {
                return false;
            }

            return ParseTokens(rowText).Any(token => detailColumns.Contains(token.Name));
        }

        private static string ValueKey(object? value)
        {
            if (value is null || value is DBNull)
            {
                return string.Empty;
            }

            return value is byte[] bytes ? Convert.ToBase64String(bytes) : value.ToString() ?? string.Empty;
        }
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: all tests PASS.

- [ ] **Step 6: STOP — present changes to the user for review. The user commits.**

---

### Task 3: Result-set reader + SQL Server / Oracle executors + options

**Files:**
- Create: `WordToPDF/Services/InvoiceResultSetReader.cs`
- Create: `WordToPDF/Services/IInvoiceDbExecutor.cs`
- Create: `WordToPDF/Models/DigitalInvoiceOptions.cs`
- Create: `WordToPDF.Tests/InvoiceResultSetReaderTests.cs`
- Modify: `WordToPDF/appsettings.json`

**Interfaces:**
- Consumes: `InvoiceDataSet` (Task 2).
- Produces:
  - `static Task<InvoiceDataSet> InvoiceResultSetReader.ReadAsync(DbDataReader reader)`
  - `interface IInvoiceDbExecutor { Task<InvoiceDataSet> ExecuteAsync(int invoiceId); }`
  - `SqlServerInvoiceExecutor(string connectionString, string procedureName) : IInvoiceDbExecutor`
  - `OracleInvoiceExecutor(string connectionString, string procedureName) : IInvoiceDbExecutor`
  - `class DigitalInvoiceOptions { const string SectionName = "DigitalInvoice"; string SqlServerProcedure; string OracleProcedure; string SqlServerConnectionStringName; string OracleConnectionStringName; }`

Note: the result set may contain case-duplicate aliases (the proc selects `aii.Id` and `ai.id`). `Dictionary` with `OrdinalIgnoreCase` keeps the last one. The delivered template never references `{Id}`, so this is harmless; document it to the user.

- [ ] **Step 1: Write the failing tests**

`WordToPDF.Tests/InvoiceResultSetReaderTests.cs`:

```csharp
using System.Data;
using WordToPDF.Services;
using Xunit;

namespace WordToPDF.Tests
{
    public class InvoiceResultSetReaderTests
    {
        [Fact]
        public async Task ReadAsync_MapsColumnsAndRows()
        {
            var table = new DataTable();
            table.Columns.Add("InvoiceNo", typeof(string));
            table.Columns.Add("Qty", typeof(decimal));
            table.Rows.Add("INV-1", 12.5m);
            table.Rows.Add("INV-2", 3m);

            using var reader = table.CreateDataReader();
            var data = await InvoiceResultSetReader.ReadAsync(reader);

            Assert.Equal(new[] { "InvoiceNo", "Qty" }, data.Columns);
            Assert.Equal(2, data.Rows.Count);
            Assert.Equal("INV-1", data.FirstRow!["InvoiceNo"]);
            Assert.Equal(12.5m, data.FirstRow!["Qty"]);
            Assert.Equal("INV-2", data.Rows[1]["InvoiceNo"]);
        }

        [Fact]
        public async Task ReadAsync_MapsDbNullToNull()
        {
            var table = new DataTable();
            table.Columns.Add("FbrInvoiceNo", typeof(string));
            table.Rows.Add(DBNull.Value);

            using var reader = table.CreateDataReader();
            var data = await InvoiceResultSetReader.ReadAsync(reader);

            Assert.Null(data.FirstRow!["FbrInvoiceNo"]);
        }

        [Fact]
        public async Task ReadAsync_LookupIsCaseInsensitive()
        {
            var table = new DataTable();
            table.Columns.Add("InvoiceNo", typeof(string));
            table.Rows.Add("INV-1");

            using var reader = table.CreateDataReader();
            var data = await InvoiceResultSetReader.ReadAsync(reader);

            Assert.Equal("INV-1", data.FirstRow!["invoiceno"]);
        }

        [Fact]
        public async Task ReadAsync_EmptyResult_KeepsColumns()
        {
            var table = new DataTable();
            table.Columns.Add("InvoiceNo", typeof(string));

            using var reader = table.CreateDataReader();
            var data = await InvoiceResultSetReader.ReadAsync(reader);

            Assert.Single(data.Columns);
            Assert.Empty(data.Rows);
            Assert.Null(data.FirstRow);
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: build error — `InvoiceResultSetReader` does not exist.

- [ ] **Step 3: Implement reader, executors, options**

`WordToPDF/Services/InvoiceResultSetReader.cs`:

```csharp
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
```

`WordToPDF/Services/IInvoiceDbExecutor.cs`:

```csharp
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
```

`WordToPDF/Models/DigitalInvoiceOptions.cs`:

```csharp
namespace WordToPDF.Models
{
    public sealed class DigitalInvoiceOptions
    {
        public const string SectionName = "DigitalInvoice";

        public string SqlServerProcedure { get; set; } = "SpArDigitalInvoices";

        public string OracleProcedure { get; set; } = "SP_AR_DIGITAL_INVOICES";

        public string SqlServerConnectionStringName { get; set; } = "SqlServerConnection";

        public string OracleConnectionStringName { get; set; } = "DefaultConnection";
    }
}
```

- [ ] **Step 4: Add config section**

In `WordToPDF/appsettings.json`, inside the root object (after `"ConnectionStrings"`), add:

```json
  "DigitalInvoice": {
    "SqlServerProcedure": "SpArDigitalInvoices",
    "OracleProcedure": "SP_AR_DIGITAL_INVOICES",
    "SqlServerConnectionStringName": "SqlServerConnection",
    "OracleConnectionStringName": "DefaultConnection"
  },
```

Keep JSON valid (mind the comma after the `ConnectionStrings` closing brace).

- [ ] **Step 5: Run tests to verify they pass**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: all tests PASS.

- [ ] **Step 6: STOP — present changes to the user for review. The user commits.**

---

### Task 4: OpenXML template filler

**Files:**
- Create: `WordToPDF/Services/InvoiceDocxFiller.cs`
- Create: `WordToPDF.Tests/InvoiceDocxFillerTests.cs`

**Interfaces:**
- Consumes: `InvoicePlaceholders` (`TokenPattern`, `ParseTokens`, `FormatValue`, `GetDetailColumns`, `IsDetailRow`), `InvoiceDataSet` (Tasks 1–2).
- Produces:
  - `class InvoiceFillResult { byte[] DocxBytes; List<string> UnresolvedPlaceholders; }`
  - `static InvoiceFillResult InvoiceDocxFiller.Fill(byte[] templateBytes, InvoiceDataSet data)`

- [ ] **Step 1: Write the failing tests**

`WordToPDF.Tests/InvoiceDocxFillerTests.cs`:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordToPDF.Models;
using WordToPDF.Services;
using Xunit;

namespace WordToPDF.Tests
{
    public class InvoiceDocxFillerTests
    {
        private static byte[] BuildDoc(params OpenXmlElement[] bodyChildren)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(bodyChildren));
                main.Document.Save();
            }

            return ms.ToArray();
        }

        private static Paragraph P(string text) =>
            new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

        private static Table T(params TableRow[] rows) => new(rows);

        private static TableRow R(params string[] cells) =>
            new(cells.Select(c => (OpenXmlElement)new TableCell(new Paragraph(new Run(new Text(c) { Space = SpaceProcessingModeValues.Preserve })))));

        private static string GetDocText(byte[] docx)
        {
            using var ms = new MemoryStream(docx);
            using var doc = WordprocessingDocument.Open(ms, false);
            return doc.MainDocumentPart!.Document.InnerText!;
        }

        private static List<string> GetTableRowTexts(byte[] docx)
        {
            using var ms = new MemoryStream(docx);
            using var doc = WordprocessingDocument.Open(ms, false);
            return doc.MainDocumentPart!.Document.Body!.Descendants<TableRow>()
                .Select(r => r.InnerText ?? string.Empty)
                .ToList();
        }

        private static InvoiceDataSet Data()
        {
            var data = new InvoiceDataSet
            {
                Columns = { "InvoiceNo", "Qty", "Rate", "GrandTotal" }
            };
            data.Rows.Add(new Dictionary<string, object?>
            {
                ["InvoiceNo"] = "5410", ["Qty"] = 1m, ["Rate"] = 100m, ["GrandTotal"] = 1500m
            });
            data.Rows.Add(new Dictionary<string, object?>
            {
                ["InvoiceNo"] = "5410", ["Qty"] = 2m, ["Rate"] = 200m, ["GrandTotal"] = 1500m
            });
            return data;
        }

        [Fact]
        public void Fill_ReplacesMasterPlaceholderInParagraph()
        {
            var docx = BuildDoc(P("Invoice No: {InvoiceNo}"));

            var result = InvoiceDocxFiller.Fill(docx, Data());

            Assert.Contains("Invoice No: 5410", GetDocText(result.DocxBytes));
            Assert.Empty(result.UnresolvedPlaceholders);
        }

        [Fact]
        public void Fill_ClonesDetailRowPerDataRow()
        {
            var docx = BuildDoc(T(R("Qty", "Rate"), R("{Qty}", "{Rate:N0}")));

            var result = InvoiceDocxFiller.Fill(docx, Data());

            var rows = GetTableRowTexts(result.DocxBytes);
            Assert.Equal(4, rows.Count);
            Assert.Equal("QtyRate", rows[0]);
            Assert.Equal("1100", rows[1]);
            Assert.Equal("2200", rows[2]);
            Assert.Contains("1,500", rows[3]);
        }

        [Fact]
        public void Fill_MasterTokenRowInTable_IsNotRepeated()
        {
            var docx = BuildDoc(T(R("Qty", "Rate"), R("{Qty}", "{Rate:N0}"), R("Grand Total", "{GrandTotal:N0}")));

            var result = InvoiceDocxFiller.Fill(docx, Data());

            var rows = GetTableRowTexts(result.DocxBytes);
            Assert.Equal(3, rows.Count);
            Assert.Contains("1,500", rows[2]);
        }

        [Fact]
        public void Fill_RemovesTemplateDetailRowTokens()
        {
            var docx = BuildDoc(T(R("{Qty}", "{Rate:N0}")));

            var result = InvoiceDocxFiller.Fill(docx, Data());

            Assert.DoesNotContain("{Qty}", GetDocText(result.DocxBytes));
            Assert.DoesNotContain("{Rate", GetDocText(result.DocxBytes));
        }

        [Fact]
        public void Fill_UnknownPlaceholder_LeftInPlaceAndReported()
        {
            var docx = BuildDoc(P("Agent: {NoSuchField}"));

            var result = InvoiceDocxFiller.Fill(docx, Data());

            Assert.Contains("Agent: {NoSuchField}", GetDocText(result.DocxBytes));
            Assert.Contains("NoSuchField", result.UnresolvedPlaceholders);
        }

        [Fact]
        public void Fill_FormatsDatesWithSpecifier()
        {
            var data = new InvoiceDataSet { Columns = { "InvoiceDate" } };
            data.Rows.Add(new Dictionary<string, object?> { ["InvoiceDate"] = new DateTime(2026, 9, 14) });

            var docx = BuildDoc(P("Date: {InvoiceDate:yyyy-MM-dd}"));
            var result = InvoiceDocxFiller.Fill(docx, data);

            Assert.Contains("Date: 2026-09-14", GetDocText(result.DocxBytes));
        }
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: build error — `InvoiceDocxFiller` does not exist.

- [ ] **Step 3: Implement the filler**

`WordToPDF/Services/InvoiceDocxFiller.cs`:

```csharp
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordToPDF.Models;

namespace WordToPDF.Services
{
    public sealed class InvoiceFillResult
    {
        public byte[] DocxBytes { get; init; } = Array.Empty<byte>();

        public List<string> UnresolvedPlaceholders { get; init; } = new();
    }

    public static class InvoiceDocxFiller
    {
        public static InvoiceFillResult Fill(byte[] templateBytes, InvoiceDataSet data)
        {
            var detailColumns = InvoicePlaceholders.GetDetailColumns(data);
            var firstRow = data.FirstRow ?? new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            var unresolved = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

            using var working = new MemoryStream();
            working.Write(templateBytes, 0, templateBytes.Length);

            using (var doc = WordprocessingDocument.Open(working, true))
            {
                var main = doc.MainDocumentPart!;

                foreach (var table in main.Document.Body!.Descendants<Table>().ToList())
                {
                    ProcessTable(table, data, detailColumns, unresolved);
                }

                ReplaceFirstRowValues(main.Document.Body, firstRow, unresolved);
                foreach (var headerPart in main.HeaderParts)
                {
                    ReplaceFirstRowValues(headerPart.Header, firstRow, unresolved);
                }

                foreach (var footerPart in main.FooterParts)
                {
                    ReplaceFirstRowValues(footerPart.Footer, firstRow, unresolved);
                }

                main.Document.Save();
            }

            return new InvoiceFillResult
            {
                DocxBytes = working.ToArray(),
                UnresolvedPlaceholders = unresolved.ToList()
            };
        }

        private static void ProcessTable(
            Table table,
            InvoiceDataSet data,
            HashSet<string> detailColumns,
            ISet<string> unresolved)
        {
            foreach (var templateRow in table.Elements<TableRow>().ToList())
            {
                if (!InvoicePlaceholders.IsDetailRow(templateRow.InnerText ?? string.Empty, detailColumns))
                {
                    continue;
                }

                var insertIndex = table.Elements<TableRow>().ToList().IndexOf(templateRow);
                foreach (var dataRow in data.Rows)
                {
                    var clone = (TableRow)templateRow.CloneNode(true);
                    foreach (var paragraph in clone.Descendants<Paragraph>())
                    {
                        ReplaceTokensInParagraph(
                            paragraph,
                            name => dataRow.TryGetValue(name, out var value) ? (true, value) : (false, null),
                            unresolved);
                    }

                    table.InsertAt(clone, insertIndex);
                    insertIndex++;
                }

                templateRow.Remove();
            }
        }

        private static void ReplaceFirstRowValues(
            OpenXmlElement element,
            Dictionary<string, object?> firstRow,
            ISet<string> unresolved)
        {
            foreach (var paragraph in element.Descendants<Paragraph>().ToList())
            {
                ReplaceTokensInParagraph(
                    paragraph,
                    name => firstRow.TryGetValue(name, out var value) ? (true, value) : (false, null),
                    unresolved);
            }
        }

        private static void ReplaceTokensInParagraph(
            Paragraph paragraph,
            Func<string, (bool Found, object? Value)> resolver,
            ISet<string> unresolved)
        {
            var runs = paragraph.Elements<Run>().ToList();
            if (runs.Count == 0)
            {
                return;
            }

            var combined = new StringBuilder();
            foreach (var run in runs)
            {
                var text = run.GetFirstChild<Text>();
                if (text != null)
                {
                    combined.Append(text.Text);
                }
            }

            var fullText = combined.ToString();
            if (InvoicePlaceholders.ParseTokens(fullText).Count == 0)
            {
                return;
            }

            var replaced = InvoicePlaceholders.TokenPattern.Replace(fullText, match =>
            {
                var name = match.Groups["name"].Value.Trim();
                var format = match.Groups["format"].Success ? match.Groups["format"].Value.Trim() : null;
                var (found, value) = resolver(name);
                if (!found)
                {
                    unresolved.Add(name);
                    return match.Value;
                }

                return InvoicePlaceholders.FormatValue(value, format);
            });

            var targetRun = runs.First(r => r.GetFirstChild<Text>() != null);
            var runProperties = targetRun.RunProperties?.CloneNode(true) as RunProperties;

            foreach (var run in runs)
            {
                var text = run.GetFirstChild<Text>();
                if (run == targetRun || text == null)
                {
                    continue;
                }

                run.RemoveChild(text);
            }

            var targetText = targetRun.GetFirstChild<Text>()!;
            targetText.Text = replaced;
            targetText.Space = SpaceProcessingModeValues.Preserve;
            if (runProperties != null && targetRun.RunProperties == null)
            {
                targetRun.RunProperties = runProperties;
            }
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: all tests PASS.

- [ ] **Step 5: STOP — present changes to the user for review. The user commits.**

---

### Task 5: Digital invoice service + controller + DI wiring

**Files:**
- Create: `WordToPDF/Services/DigitalInvoiceService.cs`
- Create: `WordToPDF/Controllers/DigitalInvoiceController.cs`
- Modify: `WordToPDF/Program.cs`

**Interfaces:**
- Consumes: `IInvoiceDbExecutor` implementations, `InvoiceDocxFiller.Fill`, `InvoiceDataSet`, `DigitalInvoiceOptions`, `ITemplateService.SaveOrUpdateTemplateAsync(string, Stream, bool)` / `GetTemplateStreamAsync(int)`, `PDFService.ResolveSofficePath()`.
- Produces:
  - `interface IDigitalInvoiceService { Task<InvoiceDataSet> GetInvoiceDataAsync(int invoiceId, string dbType); }`
  - `static string? DigitalInvoiceService.NormalizeDbType(string? dbType)` → `"sqlserver" | "oracle" | null`
  - Endpoints: `GET /DigitalInvoice/placeholders`, `POST /DigitalInvoice/templates`, `POST /DigitalInvoice/generate`

- [ ] **Step 1: Implement the service**

`WordToPDF/Services/DigitalInvoiceService.cs`:

```csharp
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
```

- [ ] **Step 2: Implement the controller**

`WordToPDF/Controllers/DigitalInvoiceController.cs`:

```csharp
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using WordToPDF.Models;
using WordToPDF.Services;

namespace WordToPDF.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class DigitalInvoiceController : ControllerBase
    {
        private readonly IDigitalInvoiceService _invoiceService;
        private readonly ITemplateService _templateService;

        public DigitalInvoiceController(IDigitalInvoiceService invoiceService, ITemplateService templateService)
        {
            _invoiceService = invoiceService;
            _templateService = templateService;
        }

        [HttpGet("placeholders")]
        public async Task<IActionResult> GetPlaceholders(
            [FromQuery] string dbType = "sqlserver",
            [FromQuery] int sampleId = 0)
        {
            try
            {
                var data = await _invoiceService.GetInvoiceDataAsync(sampleId, dbType);
                return Ok(new
                {
                    dbType = DigitalInvoiceService.NormalizeDbType(dbType),
                    count = data.Columns.Count,
                    placeholders = data.Columns
                });
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("templates")]
        public async Task<IActionResult> UploadTemplate(
            [FromForm] string templateName,
            [FromForm] bool overwrite,
            IFormFile file)
        {
            if (string.IsNullOrWhiteSpace(templateName))
            {
                return BadRequest("templateName is required.");
            }

            if (file is null || file.Length == 0)
            {
                return BadRequest("No Word (.docx) file uploaded.");
            }

            if (!Path.GetExtension(file.FileName).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("Only .docx files are supported.");
            }

            try
            {
                var template = await _templateService.SaveOrUpdateTemplateAsync(templateName, file.OpenReadStream(), overwrite);
                return Ok(new { template.Id, template.TemplateName, template.FilePath });
            }
            catch (InvalidOperationException ex)
            {
                return Conflict(ex.Message);
            }
        }

        [HttpPost("generate")]
        public async Task<IActionResult> Generate(
            [FromForm] int invoiceId,
            [FromForm] int templateId,
            [FromForm] string dbType = "sqlserver")
        {
            if (invoiceId <= 0)
            {
                return BadRequest("invoiceId is required.");
            }

            if (templateId <= 0)
            {
                return BadRequest("templateId is required.");
            }

            byte[] templateBytes;
            using (var templateStream = await _templateService.GetTemplateStreamAsync(templateId))
            {
                if (templateStream == null)
                {
                    return NotFound($"Template with ID {templateId} not found or file does not exist.");
                }

                using var ms = new MemoryStream();
                await templateStream.CopyToAsync(ms);
                templateBytes = ms.ToArray();
            }

            InvoiceDataSet data;
            try
            {
                data = await _invoiceService.GetInvoiceDataAsync(invoiceId, dbType);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }

            if (data.Rows.Count == 0)
            {
                return NotFound($"No data found for invoice {invoiceId}.");
            }

            InvoiceFillResult filled;
            try
            {
                filled = InvoiceDocxFiller.Fill(templateBytes, data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Failed to fill template: {ex.Message}");
            }

            string root = Path.Combine(Path.GetTempPath(), "DigitalInvoice_" + Guid.NewGuid().ToString("N"));
            string docxDir = Path.Combine(root, "docx");
            string pdfDir = Path.Combine(root, "pdf");
            Directory.CreateDirectory(docxDir);
            Directory.CreateDirectory(pdfDir);

            try
            {
                string docxPath = Path.Combine(docxDir, $"Invoice_{invoiceId}.docx");
                await System.IO.File.WriteAllBytesAsync(docxPath, filled.DocxBytes);

                string sofficePath = PDFService.ResolveSofficePath();
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sofficePath,
                    Arguments = $"--headless --convert-to pdf --outdir \"{pdfDir}\" \"{docxPath}\"",
                    WorkingDirectory = docxDir,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = new System.Diagnostics.Process { StartInfo = psi })
                {
                    process.Start();
                    string stdOut = await process.StandardOutput.ReadToEndAsync();
                    string stdErr = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();

                    if (process.ExitCode != 0)
                    {
                        return StatusCode(500,
                            $"LibreOffice conversion failed (ExitCode {process.ExitCode}).\nOUT:\n{stdOut}\nERR:\n{stdErr}");
                    }
                }

                string pdfPath = Path.Combine(pdfDir, $"Invoice_{invoiceId}.pdf");
                if (!System.IO.File.Exists(pdfPath))
                {
                    return StatusCode(500, $"PDF was not created at: {pdfPath}");
                }

                byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
                if (filled.UnresolvedPlaceholders.Count > 0)
                {
                    Response.Headers["X-Unresolved-Placeholders"] = string.Join(",", filled.UnresolvedPlaceholders);
                }

                return File(pdfBytes, "application/pdf", $"Invoice_{invoiceId}.pdf");
            }
            finally
            {
                try
                {
                    if (Directory.Exists(root))
                    {
                        Directory.Delete(root, recursive: true);
                    }
                }
                catch
                {
                }
            }
        }
    }
}
```

- [ ] **Step 3: Wire DI in Program.cs**

In `WordToPDF/Program.cs`, after the existing `builder.Services.AddScoped<...>` lines (line 36 area), add:

```csharp
builder.Services.Configure<DigitalInvoiceOptions>(
    builder.Configuration.GetSection(DigitalInvoiceOptions.SectionName));
builder.Services.AddScoped<IDigitalInvoiceService, DigitalInvoiceService>();
```

- [ ] **Step 4: Build and run full test suite**

```bash
dotnet build WordToPDF.sln
dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj
```

Expected: build succeeds, all tests PASS.

- [ ] **Step 5: Manual verification via Swagger (user runs the API)**

```bash
dotnet run --project WordToPDF/WordToPDF.csproj
```

Then in Swagger (`https://localhost:<port>/swagger`):
1. `GET /DigitalInvoice/placeholders?dbType=sqlserver` → JSON listing every alias of the current `SpArDigitalInvoices` (rows may be empty with sampleId=0; columns still listed).
2. Verify the list matches the procedure aliases (InvoiceNo, InvoiceDate, CustomerName, Qty, Rate, ...).

Expected: alias list returned; any alias added later to the proc appears here without code changes.

- [ ] **Step 6: STOP — present changes to the user for review. The user commits.**

---

### Task 6: Word template generator + end-to-end verification

**Files:**
- Create: `tools/TemplateGenerator/TemplateGenerator.csproj`
- Create: `tools/TemplateGenerator/Program.cs`
- Create (generated output): `WordTemplates/DigitalInvoiceTemplate.docx`

**Interfaces:**
- Consumes: nothing in code; produces the template the Task 5 endpoints consume.
- Produces: `WordTemplates/DigitalInvoiceTemplate.docx` with SALES TAX INVOICE layout and placeholders matching proc aliases; the items row uses detail aliases (`Qty`, `Rate`, ...), totals row uses master alias (`GrandTotal`).

- [ ] **Step 1: Create the generator project**

`tools/TemplateGenerator/TemplateGenerator.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
  </ItemGroup>

</Project>
```

- [ ] **Step 2: Write the generator program**

`tools/TemplateGenerator/Program.cs`:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

internal class Program
{
    private static void Main()
    {
        string outDir = Path.Combine(Directory.GetCurrentDirectory(), "WordTemplates");
        Directory.CreateDirectory(outDir);
        string outPath = Path.Combine(outDir, "DigitalInvoiceTemplate.docx");

        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(
                   outPath, DocumentFormat.OpenXml.Packaging.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();

            body.Append(Title("SALES TAX INVOICE"));
            body.Append(Empty());

            var meta = new Table(
                new TableRow(
                    Cell("Invoice No : {InvoiceNo}", 3200),
                    Cell("Fbr Invoice No : {FbrInvoiceNo}", 3200),
                    Cell("Date : {InvoiceDate}", 3200)));
            body.Append(meta);
            body.Append(Empty());

            var supplier = new TableCell(
                CellProps(4600),
                P("Supplier", true),
                P("Name : {CompanyName}"),
                P("Ph No & Address : {CompanyAddress}"),
                P("ST Reg No : {CompanySTN}"),
                P("NTN # : {CompanyNTN}"),
                P("Province : {SellerProvince}"));

            var buyer = new TableCell(
                CellProps(5000),
                P("Buyer", true),
                P("Name : {CustomerName}"),
                P("Address : {CustomerAddress}"),
                P("ST Reg No : {CustomerPhNo}"),
                P("NTN # : {BuyerRegistrationNo}"),
                P("Province : {BuyerProvince}"));

            body.Append(new Table(new TableRow(supplier, buyer)));
            body.Append(Empty());

            var terms = new Table(
                new TableRow(
                    Cell("Terms Of Sale : {SalesType}", 4800),
                    Cell("OGPNo : {OGPNo}", 4800)),
                new TableRow(
                    Cell("Reference No : {ReferenceNo}", 4800),
                    Cell("Customer PO No : {CustomerPono}", 4800)),
                new TableRow(
                    Cell("Remarks : {MasterRemarks}", 9600)));
            body.Append(terms);
            body.Append(Empty());

            body.Append(ItemsTable());
            body.Append(Empty());

            body.Append(P("Sales Tax : {SalesTaxAmount:N0}", bold: true));
            body.Append(P("Further Tax : {FurtherTax:N0}", bold: true));
            body.Append(P("Net Tax Inclusive Value : {GrandTotal:N0}", bold: true));
            body.Append(Empty());
            body.Append(P("This is a system-generated invoice and does not require any stamp or signature."));

            main.Document.Append(body);
            main.Document.Save();
        }

        Console.WriteLine($"Template written: {outPath}");
    }

    private static Paragraph Title(string text)
    {
        var run = new Run(
            new RunProperties(new Bold(), new FontSize { Val = "32" }),
            new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            run);
    }

    private static Paragraph P(string text, bool bold = false)
    {
        var runProperties = new RunProperties();
        if (bold)
        {
            runProperties.AppendChild(new Bold());
        }

        return new Paragraph(new Run(runProperties, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Paragraph Empty() => new(new Run(new Text("")));

    private static TableCellProperties CellProps(int width, bool header = false)
    {
        var props = new TableCellProperties(
            new TableCellWidth { Width = width.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableCellBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 }));
        if (header)
        {
            props.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Fill = "D9D9D9" });
        }

        return props;
    }

    private static TableCell Cell(string text, int width, bool bold = false, bool header = false)
    {
        return new TableCell(CellProps(width, header), P(text, bold));
    }

    private static Table ItemsTable()
    {
        var widths = new[] { 900, 2600, 1100, 900, 800, 900, 900, 1200, 800, 1200, 1200 };
        var grid = new TableGrid(widths.Select(w => new GridColumn { Width = w.ToString() }));

        string[] headers =
        {
            "Qty", "Description", "PO NO", "Lot Number", "UOM", "HS Code",
            "Rate", "Value Excl. Sales Tax (Rs)", "Sales Tax %", "Total Sales Tax (Rs)", "Value Incl. Sales Tax (Rs)"
        };

        string[] detailPlaceholders =
        {
            "{Qty:N2}", "{ItemDescription}", "{CustomerPono}", "{SRONo}", "{UOM}", "{HsCode}",
            "{Rate:N2}", "{LineTotal:N0}", "{SalesTaxPercentage:N0}", "{SalesTaxAmount:N0}", "{Total:N0}"
        };

        string[] totals =
        {
            "Total : {MasterDiscount:N0}", "", "", "", "", "", "", "{IncomeTaxAmt:N0}", "", "{FurtherTax:N0}", "{GrandTotal:N0}"
        };

        var table = new Table(
            new TableProperties(
                new TableWidth { Width = "13400", Type = TableWidthUnitValues.Dxa },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })),
            grid,
            new TableRow(headers.Select((h, i) => Cell(h, widths[i], bold: true, header: true))),
            new TableRow(detailPlaceholders.Select((p, i) => Cell(p, widths[i]))),
            new TableRow(totals.Select((t, i) => Cell(t, widths[i], bold: true))));

        return table;
    }
}
```

- [ ] **Step 3: Generate the template**

```bash
dotnet run --project tools/TemplateGenerator
```

Expected console: `Template written: D:\Wordtopdf\WordToPDF\WordTemplates\DigitalInvoiceTemplate.docx`.

Open the file in Word and visually confirm: title, meta row, supplier/buyer boxes, terms, items table with `{Qty:N2}` row between the shaded header row and the bold totals row.

- [ ] **Step 4: End-to-end verification (user runs)**

```bash
dotnet run --project WordToPDF/WordToPDF.csproj
```

1. `POST /DigitalInvoice/templates` — form fields: `templateName=DigitalInvoice`, `overwrite=true`, `file=<DigitalInvoiceTemplate.docx>` → note returned `id`.
2. `POST /DigitalInvoice/generate` — form fields: `invoiceId=<a real invoice id from ARInvoices, e.g. 5410>`, `templateId=<id from step 1>`, `dbType=sqlserver`.
3. Open the downloaded PDF and compare against the filled screenshot: values like `1,271,187`, `228,814`, `1,500,001`, `14-Sep-26` appear; the items table has one row per invoice line; no visible `{Placeholder}` text remains (except intentionally unresolved ones, reported in the `X-Unresolved-Placeholders` response header).

Note for the user: if the proc lives in a different SQL database than `EmployeeDB`, add that DB's connection string to `ConnectionStrings` and set `DigitalInvoice:SqlServerConnectionStringName` to its key.

- [ ] **Step 5: STOP — present changes to the user for review. The user commits.**

---

### Task 7: Oracle PL/SQL procedure script

**Files:**
- Create: `Database/SpArDigitalInvoicesOracle.sql`

**Interfaces:**
- Consumes: nothing.
- Produces: `SP_AR_DIGITAL_INVOICES(p_id IN NUMBER, p_result OUT SYS_REFCURSOR)` — the Oracle twin of the T-SQL proc, with quoted aliases preserving the exact same placeholder names.

- [ ] **Step 1: Write the script**

`Database/SpArDigitalInvoicesOracle.sql`:

```sql
-- Oracle twin of SpArDigitalInvoices (SQL Server).
-- Runs in the Oracle ERP schema that owns ARInvoices / InvItems / AppCompanies / ARCustomers
-- and the helper functions GETCONFIGNAME / GETAPPCONFIGNAME.
-- Quoted aliases keep the exact placeholder casing used in Word templates.
CREATE OR REPLACE PROCEDURE SP_AR_DIGITAL_INVOICES (
    p_id     IN NUMBER,
    p_result OUT SYS_REFCURSOR
)
AS
BEGIN
    OPEN p_result FOR
    SELECT
        aii.InvoiceNo                                             AS "InvoiceNo",
        aii.Id                                                    AS "InvoiceId",
        ai.id                                                     AS "ItemId",
        aii.InvoiceDate                                           AS "InvoiceDate",
        aii.InvoiceDueDate                                        AS "InvoiceDueDate",
        aii.FbrInvoiceNo                                          AS "FbrInvoiceNo",
        (SELECT ConfigValue FROM AppCompanyConfigs
          WHERE Module = 'Global' AND ConfigName = 'Number Path') AS "Path",
        aii.FilePath                                              AS "FilePath",
        aii.InvoiceType                                           AS "InvoiceType",
        TO_CHAR(aii.InvoiceDate, 'YYYY-MM')                       AS "TaxPeriod",
        ac.Name                                                   AS "CustomerName",
        ac.GSTNo                                                  AS "CustomerPhNo",
        ac.Address                                                AS "CustomerAddress",
        ac.NTNo                                                   AS "BuyerRegistrationNo",
        (SELECT Province FROM AppCities WHERE id = ac.ProvinceId) AS "BuyerProvince",
        ccc.Name                                                  AS "CompanyName",
        ccc.AddressLine1                                          AS "CompanyAddress",
        (SELECT Province FROM AppCities WHERE id = ccc.ProvinceId) AS "SellerProvince",
        ccc.NationalTaxNumber                                     AS "CompanyNTN",
        ccc.SalesTaxNumber                                        AS "CompanySTN",
        ccc.logo                                                  AS "Companylogo",
        i.Name                                                    AS "ItemDescription",
        i.Code                                                    AS "ItemCode",
        i.PartNo                                                  AS "SRONo",
        i.AlternatePartNo                                         AS "SROItemNo",
        i.HsCode                                                  AS "HsCode",
        i.Size                                                    AS "Size",
        i.PackQty                                                 AS "PackQty",
        GETCONFIGNAME(i.Unit)                                     AS "UOM",
        ai.Qty                                                    AS "Qty",
        ai.Mrp                                                    AS "Mrp",
        ai.Rate                                                   AS "Rate",
        ai.DiscountPercentage                                     AS "DiscountPercentage",
        ai.DiscountAmount                                         AS "DiscountAmount",
        aii.IncomeTaxAmt                                          AS "IncomeTaxAmt",
        ai.SalesTaxAmount                                         AS "SaleTaxAmount",
        ai.SalesTaxPercentage                                     AS "SalesTaxPercentage",
        ai.FurtherTax                                             AS "FurtherTax",
        ai.ExtraTax                                               AS "ExtraTax",
        0                                                         AS "FED",
        0                                                         AS "WithholdingTax",
        ai.Total                                                  AS "Total",
        ai.LineTotal                                              AS "LineTotal",
        aii.InvoiceDiscount                                       AS "MasterDiscount",
        aii.GrandTotal                                            AS "GrandTotal",
        aii.ReferenceNo                                           AS "ReferenceNo",
        aii.CustomerPono                                          AS "CustomerPono",
        aii.OGPNo                                                 AS "OGPNo",
        aii.Remarks                                               AS "MasterRemarks",
        ai.Remarks                                                AS "LineRemarks",
        GETAPPCONFIGNAME(aii.PaymentMode)                         AS "SalesType",
        GETAPPCONFIGNAME(aii.PaymentTermId)                       AS "TermsOfSale",
        (SELECT NVL(VALUE, 0) FROM AppCompanySetups
          WHERE name = 'QtyDecimal' AND ROWNUM = 1)               AS "QtyRound",
        (SELECT NVL(VALUE, 0) FROM AppCompanySetups
          WHERE name = 'RateDecimal' AND ROWNUM = 1)              AS "RateRound"
    FROM ARInvoices aii
    JOIN ARInvoiceItems ai ON ai.InvoiceId = aii.Id
    JOIN InvItems i        ON i.Id = ai.itemId
    JOIN AppCompanies ccc  ON i.CompanyId = ccc.Id
    JOIN ARCustomers ac    ON aii.CustomerId = ac.Id
    WHERE aii.Id = p_id
      AND aii.isdeleted = 0
      AND ai.isdeleted = 0;
END;
/
```

- [ ] **Step 2: Deliver to the user**

Explain: run this in the Oracle ERP schema once Oracle access is available; adjust helper function names (`GETCONFIGNAME`, `GETAPPCONFIGNAME`) if the Oracle schema names them differently. The API calls it via `dbType=oracle` using `DigitalInvoice:OracleProcedure` and `OracleConnectionStringName`. It cannot be executed now (no Oracle access) — it is a reviewed-by-eye translation of the T-SQL proc.

- [ ] **Step 3: STOP — present changes to the user for review. The user commits.**

---

## Verification Checklist (after all tasks)

1. `dotnet build WordToPDF.sln` — zero errors.
2. `dotnet test WordToPDF.Tests/WordToPDF.Tests.csproj` — all green.
3. `GET /DigitalInvoice/placeholders` returns the proc aliases; adding a temp alias to the proc (e.g. `GETDATE() AS NowTest`) makes it appear with no code change.
4. `POST /DigitalInvoice/generate` with the generated template returns a PDF matching the filled screenshot for a known invoice id.
5. Multi-line invoice → multiple item rows in the PDF.
6. No commit made by the agent; all changes presented to the user.
