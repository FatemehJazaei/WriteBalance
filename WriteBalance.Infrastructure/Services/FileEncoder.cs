
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;

namespace WriteBalance.Infrastructure.Services
{
    public class FileEncoder : IFileEncoder
    {
        public async Task<string> EncodeFileToBase64Async(string folderPath, string fileName)
        {
            try
            {
                string filePath = Path.Combine(folderPath, fileName);
                string mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"File is null or empty in path : {filePath}"), $"FileEncoder:EncodeFileToBase64Async --typeReport:Error");
                    throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
                }


                if (!File.Exists(filePath))
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"File not found in path : {filePath}"), $"FileEncoder:EncodeFileToBase64Async --typeReport:Error");
                    throw new FileNotFoundException("File not found.", filePath);
                }


                var bytes = await File.ReadAllBytesAsync(filePath);
                var base64 = Convert.ToBase64String(bytes);

                return $"data:{mimeType};base64,{base64}";
            }
            catch (Exception ex){

                Logger.WriteEntry(JsonConvert.SerializeObject($"{ex}"), $"FileEncoder:EncodeFileToBase64Async --typeReport:Error");
                throw;
            }

        }
    }
}
