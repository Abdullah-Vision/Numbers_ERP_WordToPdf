using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;
using WordToPDF.Models;
using WordToPDF.Services;

namespace WordToPDF.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class DigitalInvoiceController : ControllerBase
    {
        private readonly IDigitalInvoiceService _invoiceService;
        private readonly ErpTemplateStore _templateStore;

        public DigitalInvoiceController(IDigitalInvoiceService invoiceService, ErpTemplateStore templateStore)
        {
            _invoiceService = invoiceService;
            _templateStore = templateStore;
        }

        [HttpGet("placeholders")]
        public async Task<IActionResult> GetPlaceholders(
            [FromQuery] string dbType = "sqlserver",
            [FromQuery] int sampleId = 0)
        {
            try
            {
                var data = await _invoiceService.GetInvoiceDataAsync(sampleId, dbType);
                return Ok(new
                {
                    dbType = DigitalInvoiceService.NormalizeDbType(dbType),
                    count = data.Columns.Count,
                    placeholders = data.Columns
                });
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("templates")]
        public async Task<IActionResult> UploadTemplate(
            [FromForm] string templateName,
            [FromForm] bool overwrite,
            IFormFile file)
        {
            if (string.IsNullOrWhiteSpace(templateName))
            {
                return BadRequest("templateName is required.");
            }

            if (file is null || file.Length == 0)
            {
                return BadRequest("No Word (.docx) file uploaded.");
            }

            if (!Path.GetExtension(file.FileName).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("Only .docx files are supported.");
            }

            try
            {
                using var ms = new MemoryStream();
                await file.CopyToAsync(ms);
                var stored = await _templateStore.SaveAsync(templateName.Trim(), ms.ToArray(), overwrite);
                return Ok(new { stored.Id, stored.TemplateName });
            }
            catch (InvalidOperationException ex)
            {
                return Conflict(ex.Message);
            }
        }

        [HttpPost("generate")]
        public async Task<IActionResult> Generate(
            [FromForm] int invoiceId,
            [FromForm] int templateId,
            [FromForm] string dbType = "sqlserver")
        {
            if (invoiceId <= 0)
            {
                return BadRequest("invoiceId is required.");
            }

            if (templateId <= 0)
            {
                return BadRequest("templateId is required.");
            }

            byte[]? templateBytes = await _templateStore.GetBytesAsync(templateId);
            if (templateBytes == null)
            {
                return NotFound($"Template with ID {templateId} not found.");
            }

            InvoiceDataSet data;
            try
            {
                data = await _invoiceService.GetInvoiceDataAsync(invoiceId, dbType);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }

            if (data.Rows.Count == 0)
            {
                return NotFound($"No data found for invoice {invoiceId}.");
            }

            InvoiceFillResult filled;
            try
            {
                filled = InvoiceDocxFiller.Fill(templateBytes, data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Failed to fill template: {ex.Message}");
            }

            string root = Path.Combine(Path.GetTempPath(), "DigitalInvoice_" + Guid.NewGuid().ToString("N"));
            string docxDir = Path.Combine(root, "docx");
            string pdfDir = Path.Combine(root, "pdf");
            Directory.CreateDirectory(docxDir);
            Directory.CreateDirectory(pdfDir);

            try
            {
                string docxPath = Path.Combine(docxDir, $"Invoice_{invoiceId}.docx");
                await System.IO.File.WriteAllBytesAsync(docxPath, filled.DocxBytes);

                string sofficePath = PDFService.ResolveSofficePath();
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

                using (var process = new System.Diagnostics.Process { StartInfo = psi })
                {
                    process.Start();
                    string stdOut = await process.StandardOutput.ReadToEndAsync();
                    string stdErr = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();

                    if (process.ExitCode != 0)
                    {
                        return StatusCode(500,
                            $"LibreOffice conversion failed (ExitCode {process.ExitCode}).\nOUT:\n{stdOut}\nERR:\n{stdErr}");
                    }
                }

                string pdfPath = Path.Combine(pdfDir, $"Invoice_{invoiceId}.pdf");
                if (!System.IO.File.Exists(pdfPath))
                {
                    return StatusCode(500, $"PDF was not created at: {pdfPath}");
                }

                byte[] pdfBytes = await System.IO.File.ReadAllBytesAsync(pdfPath);
                if (filled.UnresolvedPlaceholders.Count > 0)
                {
                    Response.Headers["X-Unresolved-Placeholders"] = string.Join(",", filled.UnresolvedPlaceholders);
                }

                return File(pdfBytes, "application/pdf", $"Invoice_{invoiceId}.pdf");
            }
            finally
            {
                try
                {
                    if (Directory.Exists(root))
                    {
                        Directory.Delete(root, recursive: true);
                    }
                }
                catch
                {
                }
            }
        }
    }
}
