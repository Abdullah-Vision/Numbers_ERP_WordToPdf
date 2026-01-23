using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using WordToPDF.Services;
using System.IO;
using System; // For Exception

namespace WordToPDF.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class TemplateController : ControllerBase
    {
        private readonly ITemplateService _templateService;

        public TemplateController(ITemplateService templateService)
        {
            _templateService = templateService;
        }

        [HttpGet("all")]
        public async Task<IActionResult> GetAllTemplates()
        {
            var templates = await _templateService.GetAllTemplatesAsync();
            return Ok(templates);
        }

        [HttpPost("upload")]
        // Add [FromQuery] bool overwrite = false to allow overwriting via query parameter
        public async Task<IActionResult> UploadTemplate(IFormFile file, [FromForm] string templateName, [FromQuery] bool overwrite = false)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No template file uploaded.");
            }
            if (string.IsNullOrWhiteSpace(templateName))
            {
                return BadRequest("Template name is required.");
            }
            if (Path.GetExtension(file.FileName).ToLower() != ".docx")
            {
                return BadRequest("Only .docx files are supported for templates.");
            }

            try
            {
                using (var stream = file.OpenReadStream())
                {
                    // Call a new service method that handles the overwrite logic
                    var template = await _templateService.SaveOrUpdateTemplateAsync(templateName, stream, overwrite);

                    // If overwrite was false and a conflict occurred, the service will throw, caught below.
                    // If it was true or no conflict, it will return the template.
                    return Ok(template);
                }
            }
            catch (InvalidOperationException ex) // This will now typically only be caught if overwrite is false and template exists
            {
                // This indicates a conflict, which the frontend should handle by asking for overwrite confirmation
                return Conflict(ex.Message); // 409 Conflict
            }
            catch (Exception ex)
            {
                // Log the exception details in a real application
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}