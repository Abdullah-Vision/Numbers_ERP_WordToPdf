using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using WordToPDF.Services;
using System;
using System.IO.Compression;
using Org.BouncyCastle.Asn1.Cms;
using WordToPDF.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.EntityFrameworkCore;
using System.Collections.Concurrent; 

namespace WordToPDF.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class PDFConverterController : ControllerBase
    {
        private readonly IPDFService _pdfService;
        private readonly EmployeeDbContext _dbContext;
        private readonly ITemplateService _templateService;
        private readonly IServiceProvider _serviceProvider;

        public PDFConverterController(IPDFService pdfService, ITemplateService templateService, EmployeeDbContext dbContext, IServiceProvider serviceProvider)
        {
            _pdfService = pdfService;
            _dbContext = dbContext;
            _templateService = templateService;
            _serviceProvider = serviceProvider;
        }

        [HttpPost]
        [Route("generate-pdf")]
        public async Task<IActionResult> GeneratePdf([FromForm] int? applicationId, IFormFile file)
        {
            if (file is null || file.Length == 0)
                return BadRequest("No Word (.docx) file uploaded.");

            if (Path.GetExtension(file.FileName).ToLower() != ".docx")
                return BadRequest("Only .docx files are supported.");

            // 1) Read the template ONCE to memory
            byte[] templateBytes;
            using (var ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                templateBytes = ms.ToArray();
            }

            // 2) Resolve which employee IDs to process
            List<int> applicationIds = applicationId is null
                ? await _dbContext.Employees.Select(x => x.Id).ToListAsync()
                : new List<int> { applicationId.Value };

            // 3) Prepare temp folders
            string root = Path.Combine(Path.GetTempPath(), "BulkPDF_" + Guid.NewGuid().ToString("N"));
            string docxDir = Path.Combine(root, "docx");
            string pdfDir = Path.Combine(root, "pdf");
            Directory.CreateDirectory(docxDir);
            Directory.CreateDirectory(pdfDir);

            try
            {
                // 4) PARALLEL: build all DOCX in memory, then write to disk (fast)
                var throttler = new SemaphoreSlim(Math.Max(2, Environment.ProcessorCount)); // limit parallelism

                var tasks = applicationIds.Select(async id =>
                {
                    await throttler.WaitAsync();
                    try
                    {
                        using var scope = _serviceProvider.CreateScope();
                        var dbContext = scope.ServiceProvider.GetRequiredService<EmployeeDbContext>();
                        var pdfService = scope.ServiceProvider.GetRequiredService<IPDFService>();

                        using var templateStream = new MemoryStream(templateBytes, writable: false);
                        byte[] editedDocx = await pdfService.GenerateWordFromTemplate(id, templateStream);
                        string outPath = Path.Combine(docxDir, $"Emp_{id}.docx");
                        await System.IO.File.WriteAllBytesAsync(outPath, editedDocx);
                    }
                    finally
                    {
                        throttler.Release();
                    }
                });

                await Task.WhenAll(tasks);

                // 5) SINGLE LibreOffice run to convert ALL docx ? pdf
                // Use cmd.exe for wildcard expansion (*.docx)
                string sofficePath = PDFService.ResolveSofficePath(); // central place; throws if missing

                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/C \"\"{sofficePath}\" --headless --convert-to pdf *.docx --outdir \"{pdfDir}\"\"",
                    WorkingDirectory = docxDir,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var p = new System.Diagnostics.Process { StartInfo = psi })
                {
                    p.Start();
                    string stdOut = await p.StandardOutput.ReadToEndAsync();
                    string stdErr = await p.StandardError.ReadToEndAsync();
                    await p.WaitForExitAsync();

                    if (p.ExitCode != 0)
                    {
                        return StatusCode(500,
                            $"LibreOffice batch conversion failed (ExitCode {p.ExitCode}).\nOUT:\n{stdOut}\nERR:\n{stdErr}");
                    }
                }

                // 6) Create the ZIP (keep your behavior):
                // - If null applicationId => merge all PDFs to one file and add that one file into the ZIP
                // - Else => add single PDF to ZIP
                byte[] zipBytes;

                using (var zipMs = new MemoryStream())
                using (var archive = new ZipArchive(zipMs, ZipArchiveMode.Create, leaveOpen: true))
                {
                    if (applicationId is null)
                    {
                        // Merge all PDFs first
                        var pdfFiles = applicationIds
                            .Select(id => Path.Combine(pdfDir, $"Emp_{id}.pdf"))
                            .Where(System.IO.File.Exists)
                            .ToList();

                        if (pdfFiles.Count == 0)
                            return StatusCode(500, "No PDFs were generated.");

                        byte[] merged = MergePdfFilesFromDisk(pdfFiles);
                        var entry = archive.CreateEntry("AllApplicationsMerged.pdf", CompressionLevel.Optimal);
                        using var es = entry.Open();
                        await es.WriteAsync(merged, 0, merged.Length);
                    }
                    else
                    {
                        // Single PDF ? zip
                        string singlePdf = Path.Combine(pdfDir, $"Emp_{applicationId.Value}.pdf");
                        if (!System.IO.File.Exists(singlePdf))
                            return StatusCode(500, "PDF was not generated for the requested application.");

                        var entry = archive.CreateEntry($"StudentApplication_{applicationId.Value}.pdf", CompressionLevel.Optimal);
                        using var es = entry.Open();
                        await using var fs = System.IO.File.OpenRead(singlePdf);
                        await fs.CopyToAsync(es);
                    }
                    archive.Dispose();
                    zipBytes = zipMs.ToArray();
                }

                return File(zipBytes, "application/zip", "StudentApplications.zip");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
            finally
            {
                // Cleanup (best-effort)
                try { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); } catch { /* ignore */ }
            }
        }
        private static byte[] MergePdfFilesFromDisk(IEnumerable<string> pdfPaths)
        {
            using var outMs = new MemoryStream();
            using var doc = new Document();
            using var copy = new PdfSmartCopy(doc, outMs);
            doc.Open();

            foreach (var path in pdfPaths)
            {
                using var reader = new PdfReader(path);
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    var page = copy.GetImportedPage(reader, i);
                    copy.AddPage(page);
                }
            }

            doc.Close();
            return outMs.ToArray();
        }


        [HttpGet("all-employees")]
        public async Task<IActionResult> GetAllEmployees()
        {
            var templates = await _pdfService.GetAllEmployeesAsync();
            return Ok(templates);
        }

        // --- NEW ENDPOINT FOR TEMPLATE-BASED PDF GENERATION ---
        [HttpPost]
        [Route("generate-pdf-from-template")]
        public async Task<IActionResult> GeneratePdfFromTemplate([FromForm] int? applicationId, [FromForm] int templateId)
        {
            if (templateId <= 0)
            {
                return BadRequest("Invalid template ID provided.");
            }

            // 1) Load template bytes from DB (via service)
            byte[] templateBytes;
            using (var templateStream = await _templateService.GetTemplateStreamAsync(templateId))
            {
                if (templateStream == null)
                    return NotFound($"Template with ID {templateId} not found or file does not exist.");

                using var ms = new MemoryStream();
                await templateStream.CopyToAsync(ms);
                templateBytes = ms.ToArray();
            }

            // 2) Resolve application IDs (single or all)
            List<int> applicationIds = applicationId is null
                ? await _dbContext.Employees.Select(x => x.Id).ToListAsync()
                : new List<int> { applicationId.Value };

            // 3) Prepare temp folders
            string root = Path.Combine(Path.GetTempPath(), "BulkPDF_" + Guid.NewGuid().ToString("N"));
            string docxDir = Path.Combine(root, "docx");
            string pdfDir = Path.Combine(root, "pdf");
            Directory.CreateDirectory(docxDir);
            Directory.CreateDirectory(pdfDir);

            try
            {
                // 4) Parallel DOCX generation
                var throttler = new SemaphoreSlim(Math.Max(2, Environment.ProcessorCount));

                var tasks = applicationIds.Select(async id =>
                {
                    await throttler.WaitAsync();
                    try
                    {
                        using var scope = _serviceProvider.CreateScope();
                        var dbContext = scope.ServiceProvider.GetRequiredService<EmployeeDbContext>();
                        var pdfService = scope.ServiceProvider.GetRequiredService<IPDFService>();

                        using var templateStream = new MemoryStream(templateBytes, writable: false);
                        byte[] editedDocx = await pdfService.GenerateWordFromTemplate(id, templateStream);
                        string outPath = Path.Combine(docxDir, $"Emp_{id}.docx");
                        await System.IO.File.WriteAllBytesAsync(outPath, editedDocx);
                    }
                    finally
                    {
                        throttler.Release();
                    }
                });

                await Task.WhenAll(tasks);

                // 5) Convert all DOCX ? PDF using LibreOffice batch
                string sofficePath = PDFService.ResolveSofficePath();
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/C \"\"{sofficePath}\" --headless --convert-to pdf *.docx --outdir \"{pdfDir}\"\"",
                    WorkingDirectory = docxDir,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var p = new System.Diagnostics.Process { StartInfo = psi })
                {
                    p.Start();
                    string stdOut = await p.StandardOutput.ReadToEndAsync();
                    string stdErr = await p.StandardError.ReadToEndAsync();
                    await p.WaitForExitAsync();

                    if (p.ExitCode != 0)
                    {
                        return StatusCode(500,
                            $"LibreOffice batch conversion failed (ExitCode {p.ExitCode}).\nOUT:\n{stdOut}\nERR:\n{stdErr}");
                    }
                }

                // 6) Create ZIP (same logic as first method)
                byte[] zipBytes;

                using (var zipMs = new MemoryStream())
                using (var archive = new ZipArchive(zipMs, ZipArchiveMode.Create, leaveOpen: true))
                {
                    if (applicationId is null)
                    {
                        var pdfFiles = applicationIds
                            .Select(id => Path.Combine(pdfDir, $"Emp_{id}.pdf"))
                            .Where(System.IO.File.Exists)
                            .ToList();

                        if (pdfFiles.Count == 0)
                            return StatusCode(500, "No PDFs were generated.");

                        byte[] merged = MergePdfFilesFromDisk(pdfFiles);
                        var entry = archive.CreateEntry("AllApplicationsMerged.pdf", CompressionLevel.Optimal);
                        using var es = entry.Open();
                        await es.WriteAsync(merged, 0, merged.Length);
                    }
                    else
                    {
                        string singlePdf = Path.Combine(pdfDir, $"Emp_{applicationId.Value}.pdf");
                        if (!System.IO.File.Exists(singlePdf))
                            return StatusCode(500, "PDF was not generated for the requested application.");

                        var entry = archive.CreateEntry($"StudentApplication_{applicationId.Value}.pdf", CompressionLevel.Optimal);
                        using var es = entry.Open();
                        await using var fs = System.IO.File.OpenRead(singlePdf);
                        await fs.CopyToAsync(es);
                    }

                    archive.Dispose();
                    zipBytes = zipMs.ToArray();
                }

                return File(zipBytes, "application/zip", "StudentApplications.zip");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
            finally
            {
                try { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); } catch { /* ignore */ }
            }
        }
    }
}