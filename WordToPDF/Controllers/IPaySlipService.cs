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
    public class PaySlipService : IPaySlipService
    {
        private readonly ApplicationDbContext _dbContext;

        public PaySlipService(ApplicationDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        public async Task<List<PayrollDetailDto>> GetPayrollDetailsAsync(
            string periodId, int? empId, int? company, string branch, string dept)
        {
            var results = new List<PayrollDetailDto>();
            var connection = _dbContext.Database.GetDbConnection();

            await using (connection)
            {
                await connection.OpenAsync();

                using var cmd = connection.CreateCommand();
                cmd.CommandText = "IHSAN_COTTON.GET_PAYROLL_DETAILS";
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add(new OracleParameter("P_PERIOD_ID", OracleDbType.Varchar2, periodId, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("P_EMP_ID", OracleDbType.Int32, empId ?? (object)DBNull.Value, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("P_COMPANY", OracleDbType.Int32, company ?? (object)DBNull.Value, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("P_BRANCH", OracleDbType.Varchar2, branch ?? (object)DBNull.Value, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("P_DEPT", OracleDbType.Varchar2, dept ?? (object)DBNull.Value, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("P_RESULT_SET", OracleDbType.RefCursor, ParameterDirection.Output));

                using var reader = await cmd.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    results.Add(new PayrollDetailDto
                    {
                        CompanyId = SafeGetInt32(reader, "COMPANY_ID"),
                        CompanyName = SafeGetString(reader, "COMPANY_NAME"),
                        CPhone = SafeGetString(reader, "C_PHONE"),
                        Logo = SafeGetBytes(reader, "LOGO"),
                        BranchId = SafeGetString(reader, "BRANCH_ID"),
                        BranchName = SafeGetString(reader, "BRANCH_NAME"),
                        OrgId = SafeGetString(reader, "ORG_ID"),
                        Department = SafeGetString(reader, "DEPARTMENT"),
                        PayrollId = SafeGetInt32(reader, "PAYROLL_ID"),
                        EmpId = SafeGetInt32(reader, "EMP_ID"),
                        EmpNo = SafeGetString(reader, "EMP_NO"),
                        CardId = SafeGetString(reader, "CARD_ID"),
                        EmpName = SafeGetString(reader, "EMP_NAME"),
                        CurrentPay = SafeGetDecimal(reader, "CURRENT_PAY"),
                        EmpFatherName = SafeGetString(reader, "EMP_FATHER_NAME"),
                        Doj = SafeGetString(reader, "DOJ"),
                        WorkingDays = SafeGetInt32(reader, "WORKINGDAYS"),
                        Presents = SafeGetInt32(reader, "PRESENTS"),
                        EarPay = SafeGetDecimal(reader, "EAR_PAY"),
                        RatePerDay = SafeGetDecimal(reader, "RATE_PER_DAY"),
                        Designation = SafeGetString(reader, "DESIGNATION"),
                        EmployeeType = SafeGetString(reader, "EMPLOYEE_TYPE"),
                        Nic = SafeGetString(reader, "NIC"),
                        Ntn = SafeGetString(reader, "NTN"),
                        OtAmount = SafeGetDecimal(reader, "OT_AMOUNT"),
                        Incentive = SafeGetDecimal(reader, "INCENTIVE"),
                        EarnedPay = SafeGetDecimal(reader, "EARNED_PAY"),
                        DeductionAmount = SafeGetDecimal(reader, "DEDUCTION_AMOUNT"),
                        NetSalary = SafeGetDecimal(reader, "NET_SALARY"),
                        NetPayable = SafeGetDecimal(reader, "NET_PAYABLE"),
                        AmtInWords = SafeGetString(reader, "AMT_IN_WORDS"),
                        IncName = SafeGetString(reader, "INC_NAME"),
                        DedName = SafeGetString(reader, "DED_NAME"),
                        TranType = SafeGetString(reader, "TRAN_TYPE"),
                        TranName = SafeGetString(reader, "TRAN_NAME"),
                        TranAmount = SafeGetDecimal(reader, "TRAN_AMOUNT")
                    });
                }
            }

            return results;
        }

        public async Task<byte[]> GeneratePayslipFromTemplate(
            string periodId, int? empId, int? company, string branch, string dept, Stream templateStream)
        {
            if (templateStream == null || templateStream.Length == 0)
                throw new Exception("Template stream is empty.");

            var allData = await GetPayrollDetailsAsync(periodId, empId, company, branch, dept);
            if (!allData.Any())
                throw new Exception("No payroll data found for given parameters.");

            // Group by employee (all rows with same EmpId should show in one slip)
            var firstRecord = allData.First();

            // Load template
            using var working = new MemoryStream();
            await templateStream.CopyToAsync(working);
            working.Position = 0;

            using (var doc = WordprocessingDocument.Open(working, true))
            {
                var main = doc.MainDocumentPart;

                // Basic placeholders (single values)
                var basicPlaceholders = new Dictionary<string, string>
                {
                    { "{Company_name}", firstRecord.CompanyName ?? string.Empty },
                    { "{Branch_name}", firstRecord.BranchName ?? string.Empty },
                    { "{C_Phone}", firstRecord.CPhone ?? string.Empty },
                    { "{EMP_NO}", firstRecord.EmpNo ?? string.Empty },
                    { "{CARD_ID}", firstRecord.CardId ?? string.Empty },
                    { "{EMP_NAME}", firstRecord.EmpName ?? string.Empty },
                    { "{DESIGNATION}", firstRecord.Designation ?? string.Empty },
                    { "{DEPARTMENT}", firstRecord.Department ?? string.Empty },
                    { "{EMPLOYEE_TYPE}", firstRecord.EmployeeType ?? string.Empty },
                    { "{DOJ}", firstRecord.Doj ?? string.Empty },
                    { "{NIC}", firstRecord.Nic ?? string.Empty },
                    { "{NTN}", firstRecord.Ntn ?? string.Empty },
                    { "{Current_date}", DateTime.Now.ToString("dd-MMM-yyyy") },
                    { "{PRESENTS}", firstRecord.Presents.ToString() },
                    { "{WORKING_DAYS}", firstRecord.WorkingDays.ToString() },
                    { "{RATE_PER_DA}", firstRecord.RatePerDay.ToString("N2") },
                    { "{INCENTIVE}", firstRecord.Incentive.ToString("N2") },
                    { "{OT_AMOUNT}", firstRecord.OtAmount.ToString("N2") },
                    { "{NET_SALARY}", firstRecord.NetSalary.ToString("N2") },
                    { "{NET_PAYABLE}", firstRecord.NetPayable.ToString("N2") },
                    { "{AMT_IN_WORDS}", firstRecord.AmtInWords ?? string.Empty }
                };

                // Replace basic placeholders
                ReplacePlaceholdersInBody(main.Document.Body, basicPlaceholders);

                // Handle dynamic rows for INC_NAME, EARNED_PAY, DED_NAME, DEDUCTION_AMOUNT
                InsertDynamicRowsInTable(main.Document.Body, allData);

                main.Document.Save();
            }

            return working.ToArray();
        }

        private void InsertDynamicRowsInTable(Body body, List<PayrollDetailDto> allData)
        {
            var table = body.Descendants<Table>().FirstOrDefault();
            if (table == null) return;

            // Find the row containing placeholders
            TableRow templateRow = null;
            foreach (var row in table.Elements<TableRow>())
            {
                var text = row.InnerText;
                if (text.Contains("{INC_NAME}") || text.Contains("{ADED_NAME}"))
                {
                    templateRow = row;
                    break;
                }
            }

            if (templateRow == null) return;

            // Get incomes and deductions
            var incomes = allData.Where(d => d.TranType == "INC").ToList();
            var deductions = allData.Where(d => d.TranType == "DED").ToList();

            int maxRows = Math.Max(incomes.Count, deductions.Count);

            // Create rows for each income/deduction pair
            for (int i = 0; i < maxRows; i++)
            {
                var newRow = (TableRow)templateRow.CloneNode(true);

                var incName = i < incomes.Count ? incomes[i].IncName ?? string.Empty : string.Empty;
                var earnedPay = i < incomes.Count ? incomes[i].EarnedPay.ToString("N2") : string.Empty;
                var dedName = i < deductions.Count ? deductions[i].DedName ?? string.Empty : string.Empty;
                var dedAmount = i < deductions.Count ? deductions[i].DeductionAmount.ToString("N2") : string.Empty;

                var rowPlaceholders = new Dictionary<string, string>
                {
                    { "{INC_NAME}", incName },
                    { "{EARNED_PAY}", earnedPay },
                    { "{ADED_NAME}", dedName },
                    { "{ADEDUCTION_AMOUNT}", dedAmount }
                };

                ReplacePlaceholdersInRow(newRow, rowPlaceholders);

                // Insert before template row
                table.InsertBefore(newRow, templateRow);
            }

            // Remove template row
            templateRow.Remove();
        }

        private void ReplacePlaceholdersInRow(TableRow row, Dictionary<string, string> placeholders)
        {
            foreach (var cell in row.Elements<TableCell>())
            {
                foreach (var para in cell.Elements<Paragraph>())
                {
                    ProcessParagraphForPlaceholders(para, placeholders);
                }
            }
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

        // Safe helper methods for data reading
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

    public interface IPaySlipService
    {
        Task<byte[]> GeneratePayslipFromTemplate(string periodId, int? empId, int? company,
            string branch, string dept, Stream templateStream);
    }
}