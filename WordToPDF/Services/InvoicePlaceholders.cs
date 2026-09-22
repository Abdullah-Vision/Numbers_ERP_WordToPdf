using System.Globalization;
using System.Text.RegularExpressions;
using WordToPDF.Models;

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
    }
}
