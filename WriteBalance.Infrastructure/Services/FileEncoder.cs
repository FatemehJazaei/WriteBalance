
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
        public async Task<string> EncodeFileToBase64Async(MemoryStream excelStream, string folderPath, string fileName)
        {
            try
            {

                string mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var base64 = Convert.ToBase64String(excelStream.ToArray());

                return $"data:{mimeType};base64,{base64}";
            }
            catch (Exception ex){

                Logger.WriteEntry(JsonConvert.SerializeObject($"{ex}"), $"FileEncoder:EncodeFileToBase64Async --typeReport:Error");
                throw;
            }

        }
    }
}
