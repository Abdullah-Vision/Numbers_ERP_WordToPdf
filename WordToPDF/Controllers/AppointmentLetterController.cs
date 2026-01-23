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
    public class AppointmentLetterController : ControllerBase
    {
        private readonly IAppointmentLetterService _pdfService;
        private readonly ApplicationDbContext _dbContext;
        private readonly ITemplateService _templateService;
        private readonly IServiceProvider _serviceProvider;

        public AppointmentLetterController(IAppointmentLetterService pdfService, ITemplateService templateService, ApplicationDbContext dbContext, IServiceProvider serviceProvider)
        {
            _pdfService = pdfService;
            _dbContext = dbContext;
            _templateService = templateService;
            _serviceProvider = serviceProvider;
        }

        [HttpGet]
        [Route("generate-letter")]
        public async Task<IActionResult> GeneratePdfFromTemplate([FromQuery] int USER_ID, [FromQuery] string PRINT_P, [FromQuery] string LATTER_P, [FromQuery] int Person_P, [FromQuery] int templateId)
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

            // 3) Prepare temp folders
            string root = Path.Combine(Path.GetTempPath(), "SinglePDF_" + Guid.NewGuid().ToString("N"));
            string docxDir = Path.Combine(root, "docx");
            string pdfDir = Path.Combine(root, "pdf");
            Directory.CreateDirectory(docxDir);
            Directory.CreateDirectory(pdfDir);

            try
            {
                // 3) Generate DOCX from template
                using (var scope = _serviceProvider.CreateScope())
                {
                    var pdfService = scope.ServiceProvider.GetRequiredService<IAppointmentLetterService>();
                    using var templateStream = new MemoryStream(templateBytes, writable: false);

                    byte[] editedDocx = await pdfService.GenerateWordFromTemplate(Person_P, templateStream, LATTER_P, PRINT_P, USER_ID);
                    string outPath = Path.Combine(docxDir, $"Emp_{Person_P}.docx");
                    await System.IO.File.WriteAllBytesAsync(outPath, editedDocx);
                }

                // 4) Convert DOCX → PDF using LibreOffice
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
                            $"LibreOffice conversion failed (ExitCode {p.ExitCode}).\nOUT:\n{stdOut}\nERR:\n{stdErr}");
                    }
                }

                // 5) Return PDF inline (viewer mode)
                string singlePdf = Path.Combine(pdfDir, $"Emp_{Person_P}.pdf");
                if (!System.IO.File.Exists(singlePdf))
                    return StatusCode(500, "PDF was not generated for the requested application.");

                byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(singlePdf);

                // Return PDF inline so browser opens it in viewer instead of download
                Response.Headers["Content-Disposition"] = "inline; filename=Appointment_Letter.pdf";
                return File(pdfBytes, "application/pdf", enableRangeProcessing: true);
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

    }
}