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

namespace WordToPDF.Services
{
    public class PDFService : IPDFService
    {
        // SQL Server context (EmployeeDB) - Employee/Education/WordTemplate flow
        private readonly EmployeeDbContext _dbContext;

        public PDFService(EmployeeDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        public async Task<Employee> GetEmployeeWithEducationFlatAsync(int applicationId)
        {
            var flatData = await _dbContext.Set<EmployeeEducationDto>()
                .FromSqlRaw("EXEC USP_GetEmployeeWithEducationFlat @p0", applicationId)
                .ToListAsync();

            if (flatData == null || flatData.Count == 0)
                return null;

            var firstRow = flatData.First();

            var employee = new Employee
            {
                Id = firstRow.EmployeeId,
                FullName = firstRow.FullName,
                DateOfBirth = firstRow.DateOfBirth,
                Gender = firstRow.Gender,
                Address = firstRow.Address,
                City = firstRow.City,
                StateProvince = firstRow.StateProvince,
                ZipPostalCode = firstRow.ZipPostalCode,
                Country = firstRow.Country,
                PhoneNumber = firstRow.PhoneNumber,
                EmailAddress = firstRow.EmailAddress,
                EducationRecords = flatData
                    .Where(x => x.EducationId.HasValue)
                    .Select(x => new Education
                    {
                        Id = x.EducationId.Value,
                        EmployeeId = x.EmployeeId,
                        Degree = x.Degree,
                        University = x.University,
                        YearGraduated = x.YearGraduated ?? 0
                    })
                    .ToList()
            };

            return employee;
        }



        public async Task<byte[]> GenerateWordFromTemplate(int applicationId, Stream wordTemplateStream)
        {
            if (wordTemplateStream == null || wordTemplateStream.Length == 0)
                throw new Exception("Word template stream is empty or null.");

            var data = await GetEmployeeWithEducationFlatAsync(applicationId);
            if (data == null)
                throw new Exception($"Application data not found for ID {applicationId}.");

            // Load template bytes to a new MemoryStream we can modify
            using var working = new MemoryStream();
            await wordTemplateStream.CopyToAsync(working);
            working.Position = 0;

            using (var doc = WordprocessingDocument.Open(working, true))
            {
                var main = doc.MainDocumentPart;

                var placeholders = new Dictionary<string, string>
                {
                    { "{FullName}", data.FullName ?? string.Empty },
                    { "{DateOfBirth}", data.DateOfBirth?.ToString("dd MMMM yyyy") ?? string.Empty },
                    { "{Gender}", data.Gender ?? string.Empty },
                    { "{Address}", data.Address ?? string.Empty },
                    { "{City}", data.City ?? string.Empty },
                    { "{StateProvince}", data.StateProvince ?? string.Empty },
                    { "{ZipPostalCode}", data.ZipPostalCode ?? string.Empty },
                    { "{Country}", data.Country ?? string.Empty },
                    { "{PhoneNumber}", data.PhoneNumber ?? string.Empty },
                    { "{EmailAddress}", data.EmailAddress ?? string.Empty }
                };

                // Body + headers + footers replacements
                ReplacePlaceholdersInBody(main.Document.Body, placeholders);
                foreach (var hp in main.HeaderParts) ReplacePlaceholdersInBody(hp.Header, placeholders);
                foreach (var fp in main.FooterParts) ReplacePlaceholdersInBody(fp.Footer, placeholders);

                // Education tables (all instances)
                ProcessAllEducationTables(main.Document.Body, data.EducationRecords?.ToList() ?? new List<Education>());

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

        // --- NEW METHOD TO PROCESS ALL EDUCATION TABLES ---
        private void ProcessAllEducationTables(OpenXmlElement element, List<Education> educationRecords)
        {
            // Find all instances of the placeholder paragraph
            // Use ToList() to avoid issues if elements are removed during iteration
            var placeholderParagraphs = element.Descendants<Paragraph>()
                                               .Where(p => p.InnerText.Contains("<<EDUCATION_TABLE_START>>"))
                                               .ToList();

            foreach (var placeholderParagraph in placeholderParagraphs)
            {
                // The table should be the next sibling AFTER this specific placeholder paragraph
                Table educationTable = placeholderParagraph.ElementsAfter()
                                                            .OfType<Table>()
                                                            .FirstOrDefault();

                if (educationTable == null)
                {
                    Console.WriteLine($"Education table not found immediately after placeholder at: {placeholderParagraph.InnerText}. Skipping this instance.");
                    // Remove only this specific placeholder if its table is not found
                    placeholderParagraph.Remove();
                    continue;
                }

                // Find the template row within THIS specific table
                TableRow templateRow = educationTable.Descendants<TableRow>()
                                                     .FirstOrDefault(row => row.InnerText.Contains("{Degree}") &&
                                                                            row.InnerText.Contains("{University}") &&
                                                                            row.InnerText.Contains("{YearGraduated}"));

                if (templateRow == null)
                {
                    Console.WriteLine($"Template row for education table not found within table after placeholder: {placeholderParagraph.InnerText}. Skipping this instance.");
                    // Remove only this specific placeholder if its template row is not found
                    placeholderParagraph.Remove();
                    continue;
                }

                // Clone the template row's OuterXml to find it reliably later for removal
                string templateRowOuterXml = templateRow.OuterXml;

                // Store the index of the template row to insert new rows before it
                int templateRowIndex = educationTable.Elements<TableRow>().ToList().IndexOf(templateRow);

                // Populate the table with education data
                foreach (var edu in educationRecords)
                {
                    TableRow newRow = (TableRow)templateRow.CloneNode(true);

                    foreach (TableCell cell in newRow.Descendants<TableCell>())
                    {
                        foreach (Paragraph paragraphInCell in cell.Descendants<Paragraph>())
                        {
                            var rowPlaceholders = new Dictionary<string, string>
                            {
                                { "{Degree}", edu.Degree ?? string.Empty },
                                { "{University}", edu.University ?? string.Empty },
                                { "{YearGraduated}", edu.YearGraduated.ToString() ?? string.Empty }
                            };
                            ProcessParagraphForPlaceholders(paragraphInCell, rowPlaceholders);
                        }
                    }

                    educationTable.InsertAt(newRow, templateRowIndex);
                    templateRowIndex++;
                }

                // After populating, find the original template row again using its OuterXml
                TableRow rowToRemove = educationTable.Descendants<TableRow>()
                                                     .FirstOrDefault(row => row.OuterXml == templateRowOuterXml);

                if (rowToRemove != null)
                {
                    rowToRemove.Remove();
                }
                else
                {
                    Console.WriteLine("Could not find the template row for removal by OuterXml. Attempting fallback removal by InnerText.");
                    TableRow fallbackRowToRemove = educationTable.Descendants<TableRow>()
                                                                 .FirstOrDefault(row => row.InnerText.Contains("{Degree}") &&
                                                                                        row.InnerText.Contains("{University}") &&
                                                                                        row.InnerText.Contains("{YearGraduated}"));
                    if (fallbackRowToRemove != null)
                    {
                        fallbackRowToRemove.Remove();
                    }
                    else
                    {
                        Console.WriteLine("Fallback: Could not find template row by InnerText either. Template row might still be present.");
                    }
                }

                // Finally, remove THIS specific placeholder paragraph
                placeholderParagraph.Remove();
            }
        }

        public async Task<IEnumerable<Employee>> GetAllEmployeesAsync()
        {
            return await _dbContext.Employees.ToListAsync();
        }
    }

    public interface IPDFService
    {
        Task<byte[]> GenerateWordFromTemplate(int applicationId, Stream templateStream);

        Task<IEnumerable<Employee>> GetAllEmployeesAsync();
    }
}