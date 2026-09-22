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
