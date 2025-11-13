using Azure.Core;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;

namespace WriteBalance.Infrastructure.Services
{
    public class ExcelExporter : IExcelExporter
    {
        private readonly XLWorkbook _workbook;

        private readonly ILogger<ExcelExporter> _logger;
        public ExcelExporter(ILogger<ExcelExporter> logger)
        {
            _workbook = new XLWorkbook();
            _logger = logger;
        }

        public XLWorkbook GetWorkbook() => _workbook;
        public Task<MemoryStream> CreateWorkbookAsync()
            => Task.FromResult(new MemoryStream());

        public async Task SaveAsync(MemoryStream stream, string path, string fileName)
        {
            try
            {
                _logger.LogInformation("Starting SaveAsync...");
                string folderPath = Path.Combine(path, fileName);
                _workbook.SaveAs(folderPath);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to Save final excel in  (path)", path);
                throw;
            }

        }
    }
}
