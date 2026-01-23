using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using WordToPDF.Data;
using WordToPDF.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using Oracle.ManagedDataAccess.Client;
using System.Data;

namespace WordToPDF.Services
{
    public class WaterBillService : IWaterBillService
    {
        private readonly ApplicationDbContext _dbContext;

        public WaterBillService(ApplicationDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        public async Task<List<WaterInvoiceDto>> GetWaterInvoiceDetailsAsync(int invMonth, string type)
        {
            var results = new List<WaterInvoiceDto>();
            var connection = _dbContext.Database.GetDbConnection();

            await using (connection)
            {
                await connection.OpenAsync();

                using var cmd = connection.CreateCommand();
                cmd.CommandText = "zaraj_group.get_water_invoice_details";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new OracleParameter("p_inv_month", OracleDbType.Int32, invMonth, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("p_type", OracleDbType.Varchar2, type, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("p_result_set", OracleDbType.RefCursor, ParameterDirection.Output));

                using var reader = await cmd.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    results.Add(new WaterInvoiceDto
                    {
                        InvoiceNo = SafeGetString(reader, "invoice_no"),
                        InvoiceId = SafeGetInt32(reader, "invoice_id"),
                        MembershipId = SafeGetInt32(reader, "membership_id"),
                        PropertyId = SafeGetInt32(reader, "property_id"),
                        MetId = SafeGetInt32(reader, "met_id"),
                        MeterNo = SafeGetString(reader, "meter_no"),
                        Meter = SafeGetString(reader, "meter"),
                        ConnectionDate = SafeGetString(reader, "connection_date"),
                        ConnectionType = SafeGetString(reader, "connection_type"),
                        BillingMonth = SafeGetString(reader, "billing_month"),
                        IssueDate = SafeGetString(reader, "issue_date"),
                        DueDate = SafeGetString(reader, "due_date"),
                        MembershipName = SafeGetString(reader, "membership_name"),
                        Plot = SafeGetString(reader, "plot"),
                        Street = SafeGetString(reader, "street"),
                        Sector = SafeGetString(reader, "sector"),
                        ReadingDate = SafeGetString(reader, "reading_date"),
                        CurrentReading = SafeGetDecimal(reader, "current_reading"),
                        PreviousReading = SafeGetDecimal(reader, "previous_reading"),
                        UnitConsumed = SafeGetString(reader, "unit_consumed"),
                        Liters = SafeGetString(reader, "liters"),
                        NetCost = SafeGetDecimal(reader, "net_cost"),
                        CustomerId = SafeGetString(reader, "customer_id"),
                        MetStatus = SafeGetString(reader, "met_status"),
                        MetCost = SafeGetString(reader, "met_cost"),
                        MetCostReceived = SafeGetString(reader, "met_cost_received"),
                        Balance = SafeGetDecimal(reader, "balance"),
                        InstallmentAmount = SafeGetDecimal(reader, "installment_amount"),
                        BillAdjustment = SafeGetString(reader, "bill_adjustment"),
                        InvoiceAdjusted = SafeGetDecimal(reader, "invoice_adjusted"),
                        LatePaymentSurcharge = SafeGetDecimal(reader, "late_payment_surcharge"),
                        WaterCharges = SafeGetString(reader, "water_charges"),
                        UnitRateM3 = SafeGetString(reader, "unit_rate_m3"),
                        CurrentBill = SafeGetString(reader, "current_bill"),
                        MaintenanceCharges = SafeGetString(reader, "maintenance_charges"),
                        MiscCharges = SafeGetString(reader, "misc_charges"),
                        OtherCharges = SafeGetString(reader, "other_charges"),
                        Subsidies = SafeGetString(reader, "subsidies"),
                        WithoutSurcharge = SafeGetDecimal(reader, "without_surcharge"),
                        Tot = SafeGetDecimal(reader, "tot"),
                        Lps = SafeGetDecimal(reader, "lps"),
                        Arrears = SafeGetString(reader, "arrears"),
                        Total = SafeGetDecimal(reader, "total"),
                        ForAmtInWords = SafeGetDecimal(reader, "for_amt_in_words"),
                        AmtInWords = SafeGetString(reader, "amt_in_words"),
                        PrevMonth = SafeGetString(reader, "prev_month"),
                        PrevUnitConsumed = SafeGetDecimal(reader, "prev_unit_consumed"),
                        PrevADAmount = SafeGetDecimal(reader, "prev_a_d_amount"),
                        PrevPaidAmount = SafeGetDecimal(reader, "prev_paid_amount"),
                        Attachment = SafeGetBytes(reader, "attachment")
                    });
                }
            }

            return results;
        }

        public async Task<byte[]> GenerateWaterBillFromTemplate(int invMonth, string type, Stream templateStream)
        {
            if (templateStream == null || templateStream.Length == 0)
                throw new Exception("Template stream is empty.");

            var allData = await GetWaterInvoiceDetailsAsync(invMonth, type);
            if (!allData.Any())
                throw new Exception($"No water bill data found for month {invMonth} and type {type}.");

            var firstRecord = allData.First();

            // Load template
            using var working = new MemoryStream();
            await templateStream.CopyToAsync(working);
            working.Position = 0;

            using (var doc = WordprocessingDocument.Open(working, true))
            {
                var main = doc.MainDocumentPart;

                // All placeholders from the template
                var placeholders = new Dictionary<string, string>
                {
                    { "{invoice_no}", firstRecord.InvoiceNo ?? string.Empty },
                    { "{meter_no}", firstRecord.MeterNo ?? string.Empty },
                    { "{connection_date}", firstRecord.ConnectionDate ?? string.Empty },
                    { "{connection_type}", firstRecord.ConnectionType ?? string.Empty },
                    { "{billing_month}", firstRecord.BillingMonth ?? string.Empty },
                    { "{issue_date}", firstRecord.IssueDate ?? string.Empty },
                    { "{due_date}", firstRecord.DueDate ?? string.Empty },
                    { "{membership_name}", firstRecord.MembershipName ?? string.Empty },
                    { "{plot}", firstRecord.Plot ?? string.Empty },
                    { "{street}", firstRecord.Street ?? string.Empty },
                    { "{sector}", firstRecord.Sector ?? string.Empty },
                    { "{reading_date}", firstRecord.ReadingDate ?? string.Empty },
                    { "{current_reading}", firstRecord.CurrentReading.ToString("N0") },
                    { "{previous_reading}", firstRecord.PreviousReading.ToString("N0") },
                    { "{unit_consumed}", firstRecord.UnitConsumed ?? string.Empty },
                    { "{liters}", firstRecord.Liters ?? string.Empty },
                    { "{net_cost}", firstRecord.NetCost.ToString("N2") },
                    { "{customer_id}", firstRecord.CustomerId ?? string.Empty },
                    { "{met_status}", firstRecord.MetStatus ?? string.Empty },
                    { "{met_cost}", firstRecord.MetCost ?? string.Empty },
                    { "{met_cost_received}", firstRecord.MetCostReceived ?? string.Empty },
                    { "{balance}", firstRecord.Balance.ToString("N2") },
                    { "{installment_amount}", firstRecord.InstallmentAmount.ToString("N2") },
                    { "{bill_adjustment}", firstRecord.BillAdjustment ?? string.Empty },
                    { "{late_payment_surcharge}", firstRecord.LatePaymentSurcharge.ToString("N2") },
                    { "{water_charges}", firstRecord.WaterCharges ?? string.Empty },
                    { "{unit_rate_m3}", firstRecord.UnitRateM3 ?? string.Empty },
                    { "{current_bill}", firstRecord.CurrentBill ?? string.Empty },
                    { "{maintenance_charges}", firstRecord.MaintenanceCharges ?? string.Empty },
                    { "{misc_charges}", firstRecord.MiscCharges ?? string.Empty },
                    { "{other_charges}", firstRecord.OtherCharges ?? string.Empty },
                    { "{subsidies}", firstRecord.Subsidies ?? string.Empty },
                    { "{without_surcharge}", firstRecord.WithoutSurcharge.ToString("N2") },
                    { "{tot}", firstRecord.Tot.ToString("N2") },
                    { "{lps}", firstRecord.Lps.ToString("N2") },
                    { "{arrears}", firstRecord.Arrears ?? string.Empty },
                    { "{total}", firstRecord.Total.ToString("N2") },
                    { "{amt_in_words}", firstRecord.AmtInWords ?? string.Empty },
                    { "{prev_month}", firstRecord.PrevMonth ?? string.Empty },
                    { "{prev_unit_consumed}", firstRecord.PrevUnitConsumed.ToString("N0") },
                    { "{prev_a_d_amount}", firstRecord.PrevADAmount.ToString("N2") },
                    { "{prev_paid_amount}", firstRecord.PrevPaidAmount.ToString("N2") }
                };

                // Replace placeholders
                ReplacePlaceholdersInBody(main.Document.Body, placeholders);

                // Replace in headers and footers
                foreach (var hp in main.HeaderParts)
                    ReplacePlaceholdersInBody(hp.Header, placeholders);
                foreach (var fp in main.FooterParts)
                    ReplacePlaceholdersInBody(fp.Footer, placeholders);

                main.Document.Save();
            }

            return working.ToArray();
        }

        private void ReplacePlaceholdersInBody(OpenXmlElement element, Dictionary<string, string> placeholders)
        {
            foreach (var paragraph in element.Descendants<Paragraph>().ToList())
            {
                ProcessParagraphForPlaceholders(paragraph, placeholders);
            }
        }

        private void ProcessParagraphForPlaceholders(Paragraph paragraph, Dictionary<string, string> placeholders)
        {
            var runs = paragraph.Elements<Run>().ToList();
            if (!runs.Any()) return;

            StringBuilder combinedText = new StringBuilder();
            foreach (var run in runs)
            {
                var textElement = run.GetFirstChild<Text>();
                if (textElement != null)
                {
                    combinedText.Append(textElement.Text);
                }
            }

            string fullText = combinedText.ToString();
            bool hasPlaceholder = placeholders.Keys.Any(p => fullText.Contains(p));
            if (!hasPlaceholder) return;

            // Replace placeholders in combined text
            string replacedText = fullText;
            foreach (var ph in placeholders.OrderByDescending(p => p.Key.Length))
            {
                replacedText = replacedText.Replace(ph.Key, ph.Value);
            }

            // Clear runs and create single run with replaced text
            var firstRun = runs.First();
            var runProps = firstRun.RunProperties?.CloneNode(true) as RunProperties;

            paragraph.RemoveAllChildren<Run>();

            var newRun = new Run();
            if (runProps != null) newRun.AppendChild(runProps);
            newRun.AppendChild(new Text(replacedText) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(newRun);
        }

        public static string ResolveSofficePath()
        {
            var candidates = new[]
            {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };

            foreach (var c in candidates)
            {
                if (System.IO.File.Exists(c)) return c;
            }

            var env = Environment.GetEnvironmentVariable("LIBREOFFICE_SOEXEC");
            if (!string.IsNullOrWhiteSpace(env) && System.IO.File.Exists(env))
                return env;

            throw new FileNotFoundException("LibreOffice not found.");
        }

        // Safe helper methods
        private static string SafeGetString(IDataReader reader, string columnName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(columnName);
                return reader.IsDBNull(ordinal) ? string.Empty : reader.GetString(ordinal);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int SafeGetInt32(IDataReader reader, string columnName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(columnName);
                if (reader.IsDBNull(ordinal)) return 0;

                var value = reader.GetValue(ordinal);
                return Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        private static decimal SafeGetDecimal(IDataReader reader, string columnName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(columnName);
                if (reader.IsDBNull(ordinal)) return 0;

                var value = reader.GetValue(ordinal);
                return Convert.ToDecimal(value);
            }
            catch
            {
                return 0;
            }
        }

        private static byte[] SafeGetBytes(IDataReader reader, string columnName)
        {
            try
            {
                var ordinal = reader.GetOrdinal(columnName);
                if (reader.IsDBNull(ordinal)) return null;

                return (byte[])reader.GetValue(ordinal);
            }
            catch
            {
                return null;
            }
        }
    }

    public interface IWaterBillService
    {
        Task<byte[]> GenerateWaterBillFromTemplate(int invMonth, string type, Stream templateStream);
    }
}