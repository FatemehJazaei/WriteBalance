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
        private  XLWorkbook _workbookUpload;
        private  XLWorkbook _workbookReport;
        private  XLWorkbook _workbookUploadArzi;

        public ExcelExporter()
        {
            _workbookUpload = new XLWorkbook();
            _workbookReport = new XLWorkbook();
            _workbookUploadArzi = new XLWorkbook();
        }

        public XLWorkbook GetWorkbookReport()
        {
            _workbookReport = new XLWorkbook();
            return _workbookReport;
        }
        public XLWorkbook GetWorkbookUpload()
        {
            _workbookUpload = new XLWorkbook();
            return _workbookUpload;
        }
        public XLWorkbook GetWorkbookUploadArzi()
        {
            _workbookUploadArzi = new XLWorkbook(); 
            return _workbookUploadArzi;
        }
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
            finally
            {
                if (stream != null)
                {
                    stream.SetLength(0); 
                    stream.Position = 0;  
                    stream.Dispose();    
                }

                if (_workbookReport != null)
                {
                    if (_workbookReport.Worksheets.Contains("تراز خام"))
                    {
                        _workbookReport.Worksheets.Delete("تراز خام");
                    }
                    if (_workbookReport.Worksheets.Contains("تراز اکسیر ارزی"))
                    {
                        _workbookReport.Worksheets.Delete("تراز اکسیر ارزی");
                    }
                    if (_workbookReport.Worksheets.Contains("تراز اکسیر ریالی"))
                    {
                        _workbookReport.Worksheets.Delete("تراز اکسیر ریالی");
                    }
                    _workbookReport.Dispose();   
                }
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
            finally
            {               
                if (stream != null)
                {
                    stream.SetLength(0);  
                    stream.Position = 0; 
                    stream.Dispose();     
                }

                if (_workbookUpload != null)
                {
                    if (_workbookUpload.Worksheets.Contains("Data"))
                    {
                        _workbookUpload.Worksheets.Delete("Data");
                    }

                    _workbookUpload.Dispose();  
                }
            }

        }

        public async Task SaveUploadArziAsync(MemoryStream stream, string path, string fileName)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting SaveUploadArziAsync"), $"ExcelExporter: SaveUploadArziAsync --typeReport:Info");
                string folderPath = Path.Combine(path, fileName);
                _workbookUploadArzi.SaveAs(folderPath);
                await Task.CompletedTask;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"ExcelExporter: SaveUploadArziAsync --typeReport:Error");
                throw;
            }
            finally
            {
                if (stream != null)
                {
                    stream.SetLength(0);
                    stream.Position = 0;
                    stream.Dispose();
                }

                if (_workbookUploadArzi != null)
                {
                    if(_workbookUploadArzi.Worksheets.Contains("Data"))
                    {
                        _workbookUploadArzi.Worksheets.Delete("Data");
                    }
                    _workbookUploadArzi.Dispose();
                }
            }

        }
    }
}
