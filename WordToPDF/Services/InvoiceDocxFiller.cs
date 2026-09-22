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

                var childIndex = table.ChildElements.Cast<OpenXmlElement>().ToList().IndexOf(templateRow);
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

                    table.InsertAt(clone, childIndex);
                    childIndex++;
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
        }
    }
}
