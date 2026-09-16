using WordToPDF.Data;
using WordToPDF.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace WordToPDF.Services
{
    // SQL Server version of the template service.
    // Same behavior as TemplateService, but reads/writes the WordTemplate
    // table in EmployeeDB (SQL Server) via EmployeeDbContext.
    public class SqlTemplateService : ITemplateService
    {
        private readonly EmployeeDbContext _dbContext;
        private readonly string _templateStoragePath;

        public SqlTemplateService(EmployeeDbContext dbContext)
        {
            _dbContext = dbContext;
            _templateStoragePath = Path.Combine(Directory.GetCurrentDirectory(), "WordTemplates");
            Directory.CreateDirectory(_templateStoragePath); // Ensure the directory exists
        }

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
                            throw new InvalidOperationException($"Could not delete old template file '{existingTemplate.FilePath}'. It might be in use. Error: {ex.Message}");
                        }
                    }

                    // 2. Save the new file with the existing template's file path
                    fileStream.Seek(0, SeekOrigin.Begin);
                    using (var outputStream = new FileStream(existingTemplate.FilePath, FileMode.Create))
                    {
                        await fileStream.CopyToAsync(outputStream);
                    }

                    _dbContext.WordTemplates.Update(existingTemplate);
                    await _dbContext.SaveChangesAsync();

                    return existingTemplate;
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
}
