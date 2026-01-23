using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using WordToPDF.Data;
using WordToPDF.Models;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Oracle.ManagedDataAccess.Client;
using System.Data;

namespace WordToPDF.Services
{
    public class AppointmentLetterService : IAppointmentLetterService
    {
        private readonly ApplicationDbContext _dbContext;

        public AppointmentLetterService(ApplicationDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        public async Task<AppointmentLetterDto?> GetAppointmentLettersAsync(int personId, string letter, string print,int USER_ID)
        {
            AppointmentLetterDto? dto = null;

            var connection = _dbContext.Database.GetDbConnection();
            await using (connection)
            {
                await connection.OpenAsync();

                using var cmd = connection.CreateCommand();
                cmd.CommandText = "AZGARD9.get_appointment_letter";
                cmd.CommandType = CommandType.StoredProcedure;

                // Parameters
                cmd.Parameters.Add(new OracleParameter("PERSON_P", OracleDbType.Int32, personId, ParameterDirection.Input));
                //cmd.Parameters.Add(new OracleParameter("USER_ID", OracleDbType.Int32, USER_ID, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("LATTER_P", OracleDbType.Varchar2, letter, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("PRINT_P", OracleDbType.Varchar2, print, ParameterDirection.Input));
                cmd.Parameters.Add(new OracleParameter("RESULT_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output));

                using var reader = await cmd.ExecuteReaderAsync();

                if (await reader.ReadAsync()) 
                {
                    dto = new AppointmentLetterDto
                    {
                        EmpId = Convert.ToInt32(reader["EMP_ID"]),
                        LetterDate = reader["LETTER_DATE"]?.ToString(),
                        EmpNo = reader["EMP_NO"]?.ToString(),
                        EmpName = reader["EMP_NAME"]?.ToString(),
                        MartialStatus = reader["MARTIAL_STATUS"]?.ToString(),
                        LatterP = reader["LATTER_P"]?.ToString(),
                        C_Sign = reader["C_SIGN"]?.ToString(),
                        MarriageStatus = reader["MARRIAGE_STATUS"]?.ToString(),
                        Sex = reader["SEX"]?.ToString(),
                        BloodGroup = reader["BLOOD_GROUP"]?.ToString(),
                        Qualification = reader["QUALIFICATION"]?.ToString(),
                        Dob = reader["DOB"] == DBNull.Value ? null : Convert.ToDateTime(reader["DOB"]),
                        NW_Pay = reader["NW_PAY"]?.ToString(),
                        C_CurrentPay = reader["C_CURRENT_PAY"]?.ToString(),
                        GuardianRelation = reader["GUARDIAN_RELATION"]?.ToString(),
                        EmpFatherName = reader["EMP_FATHER_NAME"]?.ToString(),
                        CurrentPay = reader["CURRENT_PAY"] == DBNull.Value ? null : Convert.ToDecimal(reader["CURRENT_PAY"]),
                        PermaAddress = reader["PERMA_ADDRESS"]?.ToString(),
                        ReportsTo = reader["REPORTS_TO"]?.ToString(),
                        Cnic = reader["CNIC"]?.ToString(),
                        Shift = reader["SHIFT"]?.ToString(),
                        DobFormatted = reader["DOB_FORMATTED"]?.ToString(),
                        Division = reader["DIVISION"]?.ToString(),
                        Department = reader["DEPARTMENT"]?.ToString(),
                        Section = reader["SECTION"]?.ToString(),
                        Designation = reader["DESIGNATION"]?.ToString(),
                        Grade = reader["GRADE"]?.ToString(),
                        Doj = reader["DOJ"]?.ToString(),
                        DisplayFromDate = reader["DISPLAY_FROM_DATE"]?.ToString(),
                        DisplayToDate = reader["DISPLAY_TO_DATE"]?.ToString(),
                        ConfirmDate = reader["CONFIRM_DATE"] == DBNull.Value ? null : Convert.ToDateTime(reader["CONFIRM_DATE"]),
                        LocationName = reader["LOCATION_NAME"]?.ToString(),
                        CompanyName = reader["COMPANY_NAME"]?.ToString(),
                        C_Logo = reader["C_LOGO"] == DBNull.Value ? null : (byte[])reader["C_LOGO"],
                        C_Address_Urdu = reader["C_ADDRESS_URDU"]?.ToString(),
                        C_Address = reader["C_ADDRESS"]?.ToString()
                    };
                }
            }

            return dto;
        }



        public async Task<byte[]> GenerateWordFromTemplate(int personId, Stream wordTemplateStream, string Letter_P, string Print_P, int USER_ID)
        {
            if (wordTemplateStream == null || wordTemplateStream.Length == 0)
                throw new Exception("Word template stream is empty or null.");

            var data = await GetAppointmentLettersAsync(personId, Letter_P, Print_P, USER_ID);
            if (data == null)
                throw new Exception($"Application data not found for ID {personId}.");

            // Load template bytes to a new MemoryStream we can modify
            using var working = new MemoryStream();
            await wordTemplateStream.CopyToAsync(working);
            working.Position = 0;

            using (var doc = WordprocessingDocument.Open(working, true))
            {
                var main = doc.MainDocumentPart;

                var placeholders = new Dictionary<string, string>
                {
                    { "{{LetterDate}}", data.LetterDate ?? string.Empty },
                    { "{{EmpNo}}", data.EmpNo ?? string.Empty },
                    { "{{EmpName}}", data.EmpName ?? string.Empty },
                    { "{{EmpFatherName}}", data.EmpFatherName ?? string.Empty },
                    { "{{Sex}}", data.Sex ?? string.Empty },
                    { "{{DobFormatted}}", data.DobFormatted ?? string.Empty },
                    { "{{Qualification}}", data.Qualification ?? string.Empty },
                    { "{{BloodGroup}}", data.BloodGroup ?? string.Empty },
                    { "{{Cnic}}", data.Cnic ?? string.Empty },
                    { "{{MarriageStatus}}", data.MarriageStatus ?? string.Empty },
                    { "{{PermaAddress}}", data.PermaAddress ?? string.Empty },
                    { "{{ConfirmDate}}", data.ConfirmDate?.ToString("yyyy-MM-dd") ?? string.Empty },
                    { "{{Department}}", data.Department ?? string.Empty },
                    { "{{Shift}}", data.Shift ?? string.Empty },
                    { "{{Designation}}", data.Designation ?? string.Empty },
                    { "{{CurrentPay}}", data.CurrentPay?.ToString("N0") ?? string.Empty },
                    { "{{C_Sign}}", data.C_Sign ?? string.Empty }
                };

                // Body + headers + footers replacements
                ReplacePlaceholdersInBody(main.Document.Body, placeholders);
                foreach (var hp in main.HeaderParts) ReplacePlaceholdersInBody(hp.Header, placeholders);
                foreach (var fp in main.FooterParts) ReplacePlaceholdersInBody(fp.Footer, placeholders);



                main.Document.Save();
            }

            // Return the modified DOCX bytes
            return working.ToArray();
        }

        public static string ResolveSofficePath()
        {
            // 1) Default (typical Windows)
            var candidates = new[]
            {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };

            foreach (var c in candidates)
            {
                if (System.IO.File.Exists(c)) return c;
            }

            // 2) Env var override (optional)
            var env = Environment.GetEnvironmentVariable("LIBREOFFICE_SOEXEC");
            if (!string.IsNullOrWhiteSpace(env) && System.IO.File.Exists(env))
                return env;

            throw new FileNotFoundException("LibreOffice soffice.exe not found. Install LibreOffice or set LIBREOFFICE_SOEXEC env var.");
        }
        // --- HELPER METHODS ---

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

            StringBuilder combinedTextBuilder = new StringBuilder();
            foreach (var run in runs)
            {
                var textElement = run.GetFirstChild<Text>();
                if (textElement != null)
                {
                    if (textElement.Space?.Value != SpaceProcessingModeValues.Preserve && textElement.Text.Contains("  "))
                    {
                        textElement.Space = SpaceProcessingModeValues.Preserve;
                    }
                    combinedTextBuilder.Append(textElement.Text);
                }
            }
            string combinedParagraphText = combinedTextBuilder.ToString();

            bool hasPlaceholder = placeholders.Keys.Any(p => combinedParagraphText.Contains(p));
            if (!hasPlaceholder) return;

            var sortedPlaceholders = placeholders.OrderByDescending(p => p.Key.Length);

            List<Run> runsToAdd = new List<Run>();
            StringBuilder currentRunTextBuffer = new StringBuilder();
            RunProperties currentSequenceRunProperties = null;

            foreach (var run in runs)
            {
                var textElement = run.GetFirstChild<Text>();
                string runContent = textElement?.Text ?? string.Empty;

                if (string.IsNullOrWhiteSpace(runContent) && textElement == null)
                {
                    runsToAdd.Add((Run)run.CloneNode(true));
                    continue;
                }

                if (currentRunTextBuffer.Length == 0)
                {
                    currentSequenceRunProperties = run.RunProperties?.CloneNode(true) as RunProperties;
                }
                currentRunTextBuffer.Append(runContent);

                string bufferText = currentRunTextBuffer.ToString();
                bool replacedInThisBuffer = false;

                foreach (var ph in sortedPlaceholders)
                {
                    int idx = bufferText.IndexOf(ph.Key);
                    if (idx != -1)
                    {
                        replacedInThisBuffer = true;
                        if (idx > 0)
                        {
                            Run preRun = new Run();
                            if (currentSequenceRunProperties != null) preRun.AppendChild(currentSequenceRunProperties.CloneNode(true));
                            preRun.AppendChild(new Text(bufferText.Substring(0, idx)) { Space = SpaceProcessingModeValues.Preserve });
                            runsToAdd.Add(preRun);
                        }

                        Run replacementRun = new Run();
                        if (currentSequenceRunProperties != null) replacementRun.AppendChild(currentSequenceRunProperties.CloneNode(true));
                        replacementRun.AppendChild(new Text(ph.Value) { Space = SpaceProcessingModeValues.Preserve });
                        runsToAdd.Add(replacementRun);

                        currentRunTextBuffer.Clear();
                        currentRunTextBuffer.Append(bufferText.Substring(idx + ph.Key.Length));
                        break;
                    }
                }

                if (!replacedInThisBuffer && (run == runs.Last() || currentRunTextBuffer.Length == 0))
                {
                    if (currentRunTextBuffer.Length > 0)
                    {
                        Run remainingRun = new Run();
                        if (currentSequenceRunProperties != null) remainingRun.AppendChild(currentSequenceRunProperties.CloneNode(true));
                        remainingRun.AppendChild(new Text(currentRunTextBuffer.ToString()) { Space = SpaceProcessingModeValues.Preserve });
                        runsToAdd.Add(remainingRun);
                    }
                    currentRunTextBuffer.Clear();
                    currentSequenceRunProperties = null;
                }
            }

            var nonRunChildren = paragraph.Elements().Where(e => !(e is Run)).ToList();
            paragraph.RemoveAllChildren();

            foreach (var child in nonRunChildren)
            {
                paragraph.AppendChild(child.CloneNode(true));
            }

            foreach (var run in runsToAdd)
            {
                paragraph.AppendChild(run);
            }
        }
        public async Task<IEnumerable<Employee>> GetAllEmployeesAsync()
        {
            return await _dbContext.Employees.ToListAsync();
        }
    }

    public interface IAppointmentLetterService
    {
        Task<byte[]> GenerateWordFromTemplate(int personId, Stream templateStream, string Letter_P, string Print_P, int USER_ID);

        Task<IEnumerable<Employee>> GetAllEmployeesAsync();
    }
}