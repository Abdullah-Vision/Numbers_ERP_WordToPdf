using WordToPDF.Data;
using WordToPDF.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace WordToPDF.Services
{
    public class TemplateService : ITemplateService
    {
        private readonly ApplicationDbContext _dbContext;
        private readonly string _templateStoragePath;

        public TemplateService(ApplicationDbContext dbContext)
        {
            _dbContext = dbContext;
            _templateStoragePath = Path.Combine(Directory.GetCurrentDirectory(), "WordTemplates");
            Directory.CreateDirectory(_templateStoragePath); // Ensure the directory exists
        }

        // Keep this method if you need a strict "save only if not exists" behavior elsewhere
        public async Task<WordTemplate> SaveTemplateAsync(string templateName, Stream fileStream)
        {
            var existingTemplate = await _dbContext.WordTemplates
                                                        .FirstOrDefaultAsync(t => t.TemplateName.ToLower() == templateName.ToLower());

            if (existingTemplate != null)
            {
                throw new InvalidOperationException($"Template with name '{templateName}' already exists. Please choose a different name or update the existing one.");
            }

            string uniqueFileName = $"{Guid.NewGuid()}_{templateName}.docx";
            string filePath = Path.Combine(_templateStoragePath, uniqueFileName);

            using (var outputStream = new FileStream(filePath, FileMode.Create))
            {
                await fileStream.CopyToAsync(outputStream);
            }

            var template = new WordTemplate
            {
                TemplateName = templateName,
                FilePath = filePath
            };

            _dbContext.WordTemplates.Add(template);
            await _dbContext.SaveChangesAsync();

            return template;
        }

        // --- NEW: Save or Update Template method ---
        public async Task<WordTemplate> SaveOrUpdateTemplateAsync(string templateName, Stream fileStream, bool overwrite)
        {
            var existingTemplate = await _dbContext.WordTemplates
                                                        .FirstOrDefaultAsync(t => t.TemplateName.ToLower() == templateName.ToLower());

            if (existingTemplate != null)
            {
                if (overwrite)
                {
                    // 1. Delete the old file if it exists
                    if (File.Exists(existingTemplate.FilePath))
                    {
                        try
                        {
                            File.Delete(existingTemplate.FilePath);
                        }
                        catch (IOException ex)
                        {
                            // Handle cases where the file might be locked. 
                            // In a real app, you might log this and try again later or indicate failure.
                            throw new InvalidOperationException($"Could not delete old template file '{existingTemplate.FilePath}'. It might be in use. Error: {ex.Message}");
                        }
                    }

                    // 2. Save the new file with the existing template's file path
                    // Ensure the stream is at the beginning for copying
                    fileStream.Seek(0, SeekOrigin.Begin);
                    using (var outputStream = new FileStream(existingTemplate.FilePath, FileMode.Create))
                    {
                        await fileStream.CopyToAsync(outputStream);
                    }

                    // 3. Update any other properties if necessary (e.g., LastModifiedDate)
                    // For now, only the file content is updated, the path and name remain the same.
                    // If you added a LastModifiedDate property to WordTemplate, you'd update it here:
                    // existingTemplate.LastModifiedDate = DateTime.UtcNow;
                    _dbContext.WordTemplates.Update(existingTemplate); // Explicitly mark as modified if needed
                    await _dbContext.SaveChangesAsync();

                    return existingTemplate; // Return the updated existing template
                }
                else
                {
                    // If not overwriting and template exists, throw conflict
                    throw new InvalidOperationException($"Template with name '{templateName}' already exists. Set 'overwrite=true' to replace it.");
                }
            }
            else
            {
                // Template does not exist, proceed with saving a new one
                string uniqueFileName = $"{Guid.NewGuid()}_{templateName}.docx";
                string filePath = Path.Combine(_templateStoragePath, uniqueFileName);

                fileStream.Seek(0, SeekOrigin.Begin); // Ensure stream is at the beginning
                using (var outputStream = new FileStream(filePath, FileMode.Create))
                {
                    await fileStream.CopyToAsync(outputStream);
                }

                var newTemplate = new WordTemplate
                {
                    TemplateName = templateName,
                    FilePath = filePath
                };

                _dbContext.WordTemplates.Add(newTemplate);
                await _dbContext.SaveChangesAsync();

                return newTemplate;
            }
        }

        public async Task<Stream> GetTemplateStreamAsync(int templateId)
        {
            var template = await _dbContext.WordTemplates.FindAsync(templateId);
            if (template == null || !File.Exists(template.FilePath))
            {
                return null;
            }
            // Use FileShare.ReadWrite to allow the file to be deleted while still open
            // though the current PDFService implementation copies the stream, so it might not be strictly necessary here.
            // Still good practice for robustness.
            return new FileStream(template.FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }

        public async Task<IEnumerable<WordTemplate>> GetAllTemplatesAsync()
        {
            return await _dbContext.WordTemplates.ToListAsync();
        }

        public async Task<WordTemplate> GetTemplateByNameAsync(string templateName)
        {
            return await _dbContext.WordTemplates
                                        .FirstOrDefaultAsync(t => t.TemplateName.ToLower() == templateName.ToLower());
        }
    }
    public interface ITemplateService
    {
        Task<WordTemplate> SaveTemplateAsync(string templateName, Stream fileStream); // Keep for direct save if needed
        Task<WordTemplate> SaveOrUpdateTemplateAsync(string templateName, Stream fileStream, bool overwrite); // New method
        Task<Stream> GetTemplateStreamAsync(int templateId);
        Task<IEnumerable<WordTemplate>> GetAllTemplatesAsync();
        Task<WordTemplate> GetTemplateByNameAsync(string templateName);
    }
}