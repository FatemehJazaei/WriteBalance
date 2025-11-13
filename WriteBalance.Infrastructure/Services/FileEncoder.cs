using Microsoft.Extensions.Logging;
using Serilog.Core;
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
    public class FileEncoder : IFileEncoder
    {
        private readonly ILogger<FileEncoder> _logger;
        public FileEncoder(ILogger<FileEncoder> logger)
        {
            _logger = logger;
        }
        public async Task<string> EncodeFileToBase64Async(string folderPath, string fileName)
        {
            try
            {
                string filePath = Path.Combine(folderPath, fileName);
                string mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                if (string.IsNullOrWhiteSpace(filePath))
                    throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));

                if (!File.Exists(filePath))
                    throw new FileNotFoundException("File not found.", filePath);

                var bytes = await File.ReadAllBytesAsync(filePath);
                var base64 = Convert.ToBase64String(bytes);

                return $"data:{mimeType};base64,{base64}";
            }
            catch (Exception ex){

                _logger.LogError(ex, "Failed to EncodeFileToBase64Async");
                throw;
            }

        }
    }
}
