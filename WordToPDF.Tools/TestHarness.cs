using System.Data;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordToPDF.Models;
using WordToPDF.Services;

namespace WordToPDF.Tools
{
    internal static class TestHarness
    {
        private static int _passed;
        private static int _failed;

        internal static int Run()
        {
            _passed = 0;
            _failed = 0;

            Check("Token: simple token parses name", () =>
            {
                var token = AssertSingle(InvoicePlaceholders.ParseTokens("Invoice No: {InvoiceNo}"));
                Expect(token.Name == "InvoiceNo", $"name was '{token.Name}'");
                Expect(token.Format == null, "format should be null");
                Expect(token.Raw == "{InvoiceNo}", "raw mismatch");
            });

            Check("Token: token with format", () =>
            {
                var token = AssertSingle(InvoicePlaceholders.ParseTokens("Date: {InvoiceDate:dd-MMM-yy}"));
                Expect(token.Name == "InvoiceDate", $"name was '{token.Name}'");
                Expect(token.Format == "dd-MMM-yy", $"format was '{token.Format}'");
            });

            Check("Token: multiple tokens", () =>
            {
                var tokens = InvoicePlaceholders.ParseTokens("{InvoiceNo} due {InvoiceDate:yyyy-MM} end");
                Expect(tokens.Count == 2, $"count was {tokens.Count}");
                Expect(tokens[0].Name == "InvoiceNo", "first name mismatch");
                Expect(tokens[1].Name == "InvoiceDate", "second name mismatch");
                Expect(tokens[1].Format == "yyyy-MM", "second format mismatch");
            });

            Check("Token: name is trimmed", () =>
            {
                var token = AssertSingle(InvoicePlaceholders.ParseTokens("{ InvoiceNo }"));
                Expect(token.Name == "InvoiceNo", $"name was '{token.Name}'");
            });

            Check("Token: no tokens returns empty", () =>
            {
                Expect(InvoicePlaceholders.ParseTokens("no tokens here").Count == 0, "expected empty");
                Expect(InvoicePlaceholders.ParseTokens(null).Count == 0, "expected empty for null");
            });

            Check("Format: null returns empty", () =>
                Expect(InvoicePlaceholders.FormatValue(null, null) == string.Empty, "not empty"));

            Check("Format: DBNull returns empty", () =>
                Expect(InvoicePlaceholders.FormatValue(DBNull.Value, null) == string.Empty, "not empty"));

            Check("Format: byte array returns empty", () =>
                Expect(InvoicePlaceholders.FormatValue(new byte[] { 1, 2 }, null) == string.Empty, "not empty"));

            Check("Format: date default dd-MMM-yy", () =>
                Expect(InvoicePlaceholders.FormatValue(new DateTime(2026, 9, 14), null) == "14-Sep-26",
                    $"got '{InvoicePlaceholders.FormatValue(new DateTime(2026, 9, 14), null)}'"));

            Check("Format: date custom format", () =>
                Expect(InvoicePlaceholders.FormatValue(new DateTime(2026, 9, 14), "yyyy-MM-dd") == "2026-09-14",
                    "custom date mismatch"));

            Check("Format: decimal default N2", () =>
                Expect(InvoicePlaceholders.FormatValue(1234567.891m, null) == "1,234,567.89",
                    $"got '{InvoicePlaceholders.FormatValue(1234567.891m, null)}'"));

            Check("Format: decimal custom N0", () =>
                Expect(InvoicePlaceholders.FormatValue(1271187m, "N0") == "1,271,187",
                    $"got '{InvoicePlaceholders.FormatValue(1271187m, "N0")}'"));

            Check("Format: integer default N2", () =>
                Expect(InvoicePlaceholders.FormatValue(18, null) == "18.00", "int mismatch"));

            Check("Format: string passes through", () =>
                Expect(InvoicePlaceholders.FormatValue("abc", null) == "abc", "string mismatch"));

            Check("Detail: varying column detected", () =>
                Expect(InvoicePlaceholders.GetDetailColumns(TwoRowData()).Contains("Qty"), "Qty not detected"));

            Check("Detail: invariant columns ignored", () =>
            {
                var detail = InvoicePlaceholders.GetDetailColumns(TwoRowData());
                Expect(!detail.Contains("Rate"), "Rate should not be detail");
                Expect(!detail.Contains("CustomerName"), "CustomerName should not be detail");
            });

            Check("Detail: null vs value counts as varying", () =>
            {
                var data = new InvoiceDataSet { Columns = { "Qty" } };
                data.Rows.Add(new Dictionary<string, object?> { ["Qty"] = 1m });
                data.Rows.Add(new Dictionary<string, object?> { ["Qty"] = null });
                Expect(InvoicePlaceholders.GetDetailColumns(data).Contains("Qty"), "Qty not detected");
            });

            Check("Detail: single row produces no detail columns", () =>
            {
                var data = new InvoiceDataSet { Columns = { "Qty" } };
                data.Rows.Add(new Dictionary<string, object?> { ["Qty"] = 1m });
                Expect(InvoicePlaceholders.GetDetailColumns(data).Count == 0, "expected none");
            });

            Check("Detail: row with detail token is detail row", () =>
            {
                var detail = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Qty" };
                Expect(InvoicePlaceholders.IsDetailRow("{Qty} | {CustomerName}", detail), "should be detail");
            });

            Check("Detail: master-only row is not detail row", () =>
            {
                var detail = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Qty" };
                Expect(!InvoicePlaceholders.IsDetailRow("Total: {GrandTotal}", detail), "should not be detail");
            });

            Check("Detail: matching is case-insensitive", () =>
            {
                var detail = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Qty" };
                Expect(InvoicePlaceholders.IsDetailRow("{qty}", detail), "case mismatch");
            });

            Check("Reader: maps columns and rows", () =>
            {
                var table = new DataTable();
                table.Columns.Add("InvoiceNo", typeof(string));
                table.Columns.Add("Qty", typeof(decimal));
                table.Rows.Add("INV-1", 12.5m);
                table.Rows.Add("INV-2", 3m);

                using var reader = table.CreateDataReader();
                var data = InvoiceResultSetReader.ReadAsync(reader).GetAwaiter().GetResult();

                Expect(data.Columns.Count == 2 && data.Columns[0] == "InvoiceNo", "columns mismatch");
                Expect(data.Rows.Count == 2, "rows mismatch");
                Expect(string.Equals(data.FirstRow!["InvoiceNo"] as string, "INV-1"), "first row value mismatch");
                Expect(Convert.ToDecimal(data.FirstRow!["Qty"]) == 12.5m, "qty mismatch");
            });

            Check("Reader: DBNull maps to null", () =>
            {
                var table = new DataTable();
                table.Columns.Add("FbrInvoiceNo", typeof(string));
                table.Rows.Add(DBNull.Value);

                using var reader = table.CreateDataReader();
                var data = InvoiceResultSetReader.ReadAsync(reader).GetAwaiter().GetResult();

                Expect(data.FirstRow!["FbrInvoiceNo"] == null, "expected null");
            });

            Check("Reader: lookup is case-insensitive", () =>
            {
                var table = new DataTable();
                table.Columns.Add("InvoiceNo", typeof(string));
                table.Rows.Add("INV-1");

                using var reader = table.CreateDataReader();
                var data = InvoiceResultSetReader.ReadAsync(reader).GetAwaiter().GetResult();

                Expect(string.Equals(data.FirstRow!["invoiceno"] as string, "INV-1"), "case lookup failed");
            });

            Check("Reader: empty result keeps columns", () =>
            {
                var table = new DataTable();
                table.Columns.Add("InvoiceNo", typeof(string));

                using var reader = table.CreateDataReader();
                var data = InvoiceResultSetReader.ReadAsync(reader).GetAwaiter().GetResult();

                Expect(data.Columns.Count == 1 && data.Rows.Count == 0 && data.FirstRow == null, "empty mismatch");
            });

            Check("Filler: replaces master placeholder in paragraph", () =>
            {
                var docx = BuildDoc(P("Invoice No: {InvoiceNo}"));
                var result = InvoiceDocxFiller.Fill(docx, SampleData());

                var text = GetDocText(result.DocxBytes);
                Expect(text.Contains("Invoice No: 5410"), $"text was '{text}'");
                Expect(result.UnresolvedPlaceholders.Count == 0, "unexpected unresolved");
            });

            Check("Filler: clones detail row per data row", () =>
            {
                var docx = BuildDoc(T(R("Qty", "Rate"), R("{Qty:N0}", "{Rate:N0}")));
                var result = InvoiceDocxFiller.Fill(docx, SampleData());

                var rows = GetTableRowTexts(result.DocxBytes);
                Expect(rows.Count == 3, $"row count was {rows.Count}");
                Expect(rows[0] == "QtyRate", $"header row was '{rows[0]}'");
                Expect(rows[1] == "1100", $"first data row was '{rows[1]}'");
                Expect(rows[2] == "2200", $"second data row was '{rows[2]}'");
            });

            Check("Filler: master token row is not repeated", () =>
            {
                var docx = BuildDoc(T(
                    R("Qty", "Rate"),
                    R("{Qty}", "{Rate:N0}"),
                    R("Grand Total", "{GrandTotal:N0}")));
                var result = InvoiceDocxFiller.Fill(docx, SampleData());

                var rows = GetTableRowTexts(result.DocxBytes);
                Expect(rows.Count == 4, $"row count was {rows.Count}");
                Expect(rows[3].Contains("1,500"), $"totals row was '{rows[3]}'");
            });

            Check("Filler: removes template detail row tokens", () =>
            {
                var docx = BuildDoc(T(R("{Qty}", "{Rate:N0}")));
                var result = InvoiceDocxFiller.Fill(docx, SampleData());

                var text = GetDocText(result.DocxBytes);
                Expect(!text.Contains("{Qty}"), "token left behind");
                Expect(!text.Contains("{Rate"), "token left behind");
            });

            Check("Filler: unknown placeholder stays and is reported", () =>
            {
                var docx = BuildDoc(P("Agent: {NoSuchField}"));
                var result = InvoiceDocxFiller.Fill(docx, SampleData());

                var text = GetDocText(result.DocxBytes);
                Expect(text.Contains("Agent: {NoSuchField}"), "token was removed");
                Expect(result.UnresolvedPlaceholders.Contains("NoSuchField"), "not reported");
            });

            Check("Filler: header stays above cloned rows in real table layout", () =>
            {
                var table = new Table(
                    new TableProperties(new TableWidth { Width = "9900", Type = TableWidthUnitValues.Dxa }),
                    new TableGrid(new GridColumn { Width = "1000" }, new GridColumn { Width = "1000" }),
                    R("Qty", "Rate"),
                    R("{Qty:N0}", "{Rate:N0}"));

                var docx = BuildDoc(table);
                var result = InvoiceDocxFiller.Fill(docx, SampleData());

                var rows = GetTableRowTexts(result.DocxBytes);
                Expect(rows.Count == 3, $"row count was {rows.Count}");
                Expect(rows[0] == "QtyRate", $"HEADER NOT FIRST — row 0 was '{rows[0]}'");
                Expect(rows[1] == "1100", $"row 1 was '{rows[1]}'");
                Expect(rows[2] == "2200", $"row 2 was '{rows[2]}'");
            });

            Check("Filler: formats dates with specifier", () =>
            {
                var data = new InvoiceDataSet { Columns = { "InvoiceDate" } };
                data.Rows.Add(new Dictionary<string, object?> { ["InvoiceDate"] = new DateTime(2026, 9, 14) });

                var docx = BuildDoc(P("Date: {InvoiceDate:yyyy-MM-dd}"));
                var result = InvoiceDocxFiller.Fill(docx, data);

                Expect(GetDocText(result.DocxBytes).Contains("Date: 2026-09-14"), "date not formatted");
            });

            Check("Filler: REAL template — header first after fill", () =>
            {
                var path = Path.Combine(Directory.GetCurrentDirectory(), "WordTemplates", "DigitalInvoiceTemplate_new.docx");
                Expect(File.Exists(path), $"template not found at {path}");

                var data = new InvoiceDataSet { Columns = { "InvoiceNo", "Qty", "SaleTaxAmount", "GrandTotal" } };
                data.Rows.Add(new Dictionary<string, object?>
                {
                    ["InvoiceNo"] = "T1", ["Qty"] = 1m, ["SaleTaxAmount"] = 100m, ["GrandTotal"] = 118m
                });
                data.Rows.Add(new Dictionary<string, object?>
                {
                    ["InvoiceNo"] = "T1", ["Qty"] = 2m, ["SaleTaxAmount"] = 200m, ["GrandTotal"] = 118m
                });

                var result = InvoiceDocxFiller.Fill(File.ReadAllBytes(path), data);
                var rows = GetTableRowTexts(result.DocxBytes);

                var headerIndex = rows.FindIndex(r => r.StartsWith("Qty", StringComparison.Ordinal));
                var firstItemIndex = rows.FindIndex(r => r.Contains("1.00"));
                Expect(headerIndex >= 0, "header row missing");
                Expect(firstItemIndex > headerIndex, $"header at {headerIndex} but first item at {firstItemIndex}");
                Expect(!result.UnresolvedPlaceholders.Contains("SaleTaxAmount"), "SaleTaxAmount unresolved");
            });

            Console.WriteLine();
            Console.WriteLine($"==== RESULTS: {_passed} passed, {_failed} failed ====");
            return _failed == 0 ? 0 : 1;
        }

        private static InvoiceDataSet SampleData()
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

        private static InvoiceDataSet TwoRowData()
        {
            var data = new InvoiceDataSet { Columns = { "Qty", "Rate", "CustomerName" } };
            data.Rows.Add(new Dictionary<string, object?> { ["Qty"] = 1m, ["Rate"] = 100m, ["CustomerName"] = "Acme" });
            data.Rows.Add(new Dictionary<string, object?> { ["Qty"] = 2m, ["Rate"] = 100m, ["CustomerName"] = "Acme" });
            return data;
        }

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
            new(cells.Select(c => (OpenXmlElement)new TableCell(
                new Paragraph(new Run(new Text(c) { Space = SpaceProcessingModeValues.Preserve })))).ToArray());

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

        private static T AssertSingle<T>(IEnumerable<T> items)
        {
            var list = items.ToList();
            Expect(list.Count == 1, $"expected 1 item, got {list.Count}");
            return list[0];
        }

        private static void Expect(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        private static void Check(string name, Action assertion)
        {
            try
            {
                assertion();
                _passed++;
                Console.WriteLine($"[PASS] {name}");
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"[FAIL] {name} :: {ex.Message}");
            }
        }
    }
}
