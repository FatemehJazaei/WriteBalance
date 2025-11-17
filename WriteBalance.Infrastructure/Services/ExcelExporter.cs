using Azure.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
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
    public class ExcelExporter : IExcelExporter
    {
        private readonly XLWorkbook _workbookUpload;
        private readonly XLWorkbook _workbookReport;

        public ExcelExporter()
        {
            _workbookReport = new XLWorkbook();
            _workbookUpload = new XLWorkbook();
        }

        public XLWorkbook GetWorkbookReport() => _workbookReport;
        public XLWorkbook GetWorkbookUpload()=> _workbookUpload ;
        public Task<MemoryStream> CreateWorkbookAsync()
            => Task.FromResult(new MemoryStream());

        public async Task SaveReportAsync(MemoryStream stream, string path, string fileName)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting SaveReportAsync"), $"ExcelExporter: SaveReportAsync --typeReport:Info");
                string folderPath = Path.Combine(path, fileName);
                _workbookReport.SaveAs(folderPath);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"ExcelExporter: SaveReportAsync --typeReport:Error");
                throw;
            }

        }
        public async Task SaveUploadAsync(MemoryStream stream, string path, string fileName)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting SaveUploadAsync"), $"ExcelExporter: SaveUploadAsync --typeReport:Info");
                string folderPath = Path.Combine(path, fileName);
                _workbookUpload.SaveAs(folderPath);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"ExcelExporter: SaveUploadAsync --typeReport:Error");
                throw;
            }

        }
    }
}
