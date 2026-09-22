namespace WordToPDF.Models
{
    public sealed class InvoiceDataSet
    {
        public List<string> Columns { get; set; } = new();

        public List<Dictionary<string, object?>> Rows { get; set; } = new();

        public Dictionary<string, object?>? FirstRow => Rows.Count > 0 ? Rows[0] : null;
    }
}
