using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using WordToPDF.Services;
using System;
using WordToPDF.Data;
using System.Linq;

namespace WordToPDF.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class WaterBillController : ControllerBase
    {
        private readonly IWaterBillService _waterBillService;
        private readonly ApplicationDbContext _dbContext;
        private readonly ITemplateService _templateService;
        private readonly IServiceProvider _serviceProvider;

        public WaterBillController(IWaterBillService waterBillService, ITemplateService templateService,
            ApplicationDbContext dbContext, IServiceProvider serviceProvider)
        {
            _waterBillService = waterBillService;
            _dbContext = dbContext;
            _templateService = templateService;
            _serviceProvider = serviceProvider;
        }

        [HttpGet]
        [Route("generate-bill")]
        public async Task<IActionResult> GenerateWaterBillPdf(
            [FromQuery] int invMonth,
            [FromQuery] string type = "WB",
            [FromQuery] int templateId = 1)
        {
            if (invMonth <= 0)
            {
                return BadRequest("Invoice Month (invMonth) is required.");
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
            string root = Path.Combine(Path.GetTempPath(), "WaterBill_" + Guid.NewGuid().ToString("N"));
            string docxDir = Path.Combine(root, "docx");
            string pdfDir = Path.Combine(root, "pdf");
            Directory.CreateDirectory(docxDir);
            Directory.CreateDirectory(pdfDir);

            try
            {
                // Generate DOCX from template
                using (var scope = _serviceProvider.CreateScope())
                {
                    var waterBillService = scope.ServiceProvider.GetRequiredService<IWaterBillService>();
                    using var templateStream = new MemoryStream(templateBytes, writable: false);

                    byte[] editedDocx = await waterBillService.GenerateWaterBillFromTemplate(
                        invMonth, type, templateStream);

                    string outPath = Path.Combine(docxDir, $"WaterBill_{invMonth}.docx");
                    await System.IO.File.WriteAllBytesAsync(outPath, editedDocx);
                }

                // Check if DOCX was created
                string docxPath = Path.Combine(docxDir, $"WaterBill_{invMonth}.docx");
                if (!System.IO.File.Exists(docxPath))
                    return StatusCode(500, $"DOCX file was not created at: {docxPath}");

                // Convert DOCX → PDF using LibreOffice
                string sofficePath = WaterBillService.ResolveSofficePath();
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

                // Return PDF
                string pdfPath = Path.Combine(pdfDir, $"WaterBill_{invMonth}.pdf");

                // List all files in pdf directory for debugging
                if (!System.IO.File.Exists(pdfPath))
                {
                    var pdfFiles = Directory.GetFiles(pdfDir);
                    var fileList = pdfFiles.Length > 0 ? string.Join(", ", pdfFiles.Select(Path.GetFileName)) : "No files found";
                    return StatusCode(500, $"Expected PDF not found at: {pdfPath}\nFiles in PDF directory: {fileList}");
                }

                byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);

                // Return PDF as downloadable file
                string fileName = $"WaterBill_{invMonth}_{DateTime.Now:yyyyMMdd}.pdf";
                return File(pdfBytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                try { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); } catch { }
            }
        }
    }
}