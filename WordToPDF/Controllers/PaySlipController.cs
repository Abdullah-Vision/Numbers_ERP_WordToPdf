using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using WordToPDF.Services;
using System;
using WordToPDF.Data;
using Microsoft.EntityFrameworkCore;
using System.Linq;

namespace WordToPDF.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class PaySlipController : ControllerBase
    {
        private readonly IPaySlipService _paySlipService;
        private readonly ApplicationDbContext _dbContext;
        private readonly ITemplateService _templateService;
        private readonly IServiceProvider _serviceProvider;

        public PaySlipController(IPaySlipService paySlipService, ITemplateService templateService,
            ApplicationDbContext dbContext, IServiceProvider serviceProvider)
        {
            _paySlipService = paySlipService;
            _dbContext = dbContext;
            _templateService = templateService;
            _serviceProvider = serviceProvider;
        }

        [HttpGet]
        [Route("generate-payslip")]
        public async Task<IActionResult> GeneratePayslipPdf(
            [FromQuery] string periodId,
            [FromQuery] int? empId = null,
            [FromQuery] int? company = null,
            [FromQuery] string branch = null,
            [FromQuery] string dept = null,
            [FromQuery] int templateId = 1)
        {
            if (string.IsNullOrWhiteSpace(periodId))
            {
                return BadRequest("Period ID is required.");
            }

            if (templateId <= 0)
            {
                return BadRequest("Invalid template ID provided.");
            }

            // Load template
            byte[] templateBytes;
            using (var templateStream = await _templateService.GetTemplateStreamAsync(templateId))
            {
                if (templateStream == null)
                    return NotFound($"Template with ID {templateId} not found.");

                using var ms = new MemoryStream();
                await templateStream.CopyToAsync(ms);
                templateBytes = ms.ToArray();
            }

            // Prepare temp folders
            string root = Path.Combine(Path.GetTempPath(), "Payslip_" + Guid.NewGuid().ToString("N"));
            string docxDir = Path.Combine(root, "docx");
            string pdfDir = Path.Combine(root, "pdf");
            Directory.CreateDirectory(docxDir);
            Directory.CreateDirectory(pdfDir);

            try
            {
                // Generate DOCX from template
                using (var scope = _serviceProvider.CreateScope())
                {
                    var paySlipService = scope.ServiceProvider.GetRequiredService<IPaySlipService>();
                    using var templateStream = new MemoryStream(templateBytes, writable: false);

                    byte[] editedDocx = await paySlipService.GeneratePayslipFromTemplate(
                        periodId, empId, company, branch, dept, templateStream);

                    string outPath = Path.Combine(docxDir, $"Payslip_{periodId}.docx");
                    await System.IO.File.WriteAllBytesAsync(outPath, editedDocx);
                }

                // Check if DOCX was created
                string docxPath = Path.Combine(docxDir, $"Payslip_{periodId}.docx");
                if (!System.IO.File.Exists(docxPath))
                    return StatusCode(500, $"DOCX file was not created at: {docxPath}");

                // Convert DOCX → PDF using LibreOffice
                string sofficePath = PaySlipService.ResolveSofficePath();
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sofficePath,
                    Arguments = $"--headless --convert-to pdf --outdir \"{pdfDir}\" \"{docxPath}\"",
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

                // Return PDF inline
                string pdfPath = Path.Combine(pdfDir, $"Payslip_{periodId}.pdf");

                // List all files in pdf directory for debugging
                if (!System.IO.File.Exists(pdfPath))
                {
                    var pdfFiles = Directory.GetFiles(pdfDir);
                    var fileList = pdfFiles.Length > 0 ? string.Join(", ", pdfFiles.Select(Path.GetFileName)) : "No files found";
                    return StatusCode(500, $"Expected PDF not found at: {pdfPath}\nFiles in PDF directory: {fileList}");
                }

                byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);

                // Return PDF as downloadable file
                string fileName = $"Payslip_{periodId}_{DateTime.Now:yyyyMMdd}.pdf";
                return File(pdfBytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
            finally
            {
                try { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); } catch { }
            }
        }
    }
}