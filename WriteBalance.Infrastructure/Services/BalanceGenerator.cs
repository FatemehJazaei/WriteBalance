using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Server.Kestrel.Core.Internal.Http;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using WriteBalance.Domain.Entities;
using WriteBalance.Application.Interfaces;
using ClosedXML.Excel;
using Azure.Core;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using System.Linq.Expressions;

namespace WriteBalance.Infrastructure.Services
{
    public class BalanceGenerator : IBalanceGenerator
    {
        private readonly ILogger<BalanceGenerator> _logger;

        public BalanceGenerator(ILogger<BalanceGenerator> logger)
        {
            _logger = logger;
        }
        public async Task<MemoryStream> GenerateTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, string FolderPath)
        {
            try
            {
                _logger.LogInformation("Starting GenerateRawTablesAsync...");

                var workbook = excelExporter.GetWorkbook();
                var stream = await GenerateRawTablesAsync(financialRecords, excelExporter, workbook, FolderPath);
                stream.Position = 0;


                var worksheet = workbook.Worksheets.Add("تراز اکسیر");
                worksheet.RightToLeft = true;
                int row = 2;

                var rows = financialRecords.Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}_{x.Moeen_Code}",
                    Col2 = $"{x.Kol_Title}_{x.Moeen_Title}",
                    Col3 = x.Mande_Bed,
                    Col4 = x.Mande_Bes,
                }).ToList();

                var mergedRows = MergeDuplicateRows(rows);

                var duplicateKeys = mergedRows
                                    .GroupBy(r => r.Col1)
                                    .Where(g => g.Count() > 1)
                                    .Select(g => g.Key)
                                    .ToList();

                if (duplicateKeys.Any())
                {
                    var dupList = string.Join(", ", duplicateKeys);
                    _logger.LogWarning($"Duplicate values found in Col1: {dupList}");
                    mergedRows = MergeDuplicateRows(mergedRows);
                }

                var emptyCol2 = mergedRows.Where(r => string.IsNullOrWhiteSpace(r.Col2)).ToList();

                if (emptyCol2.Any())
                {
                    _logger.LogWarning($"Found {emptyCol2.Count} rows with empty Col2.");
                    foreach (var item in emptyCol2)
                    {
                        item.Col2 = item.Col1;
                    }
                }

                decimal totalBed = mergedRows.Sum(r => r.Col3);
                decimal totalBes = mergedRows.Sum(r => r.Col4);

                if (totalBed != totalBes)
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"مانده بدهکار {totalBed} و مانده بستانکار {totalBes} برابر نمی باشد.  " }
                        },
                    FolderPath
                    );
                }


                foreach (var item in mergedRows)
                {
                    worksheet.Cell(row, 1).Value = item.Col1;
                    worksheet.Cell(row, 2).Value = item.Col2;
                    worksheet.Cell(row, 3).Value = item.Col3;
                    worksheet.Cell(row, 4).Value = item.Col4;
                    row++;
                }

                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch(Exception ex)
            {
                _logger.LogError(ex, "Failed to GenerateTablesAsync");
                throw;
            }
        }
        private List<ExcelRow> MergeDuplicateRows(List<ExcelRow> rows)
        {
            try
            {
                var merged = rows
                                .GroupBy(r => r.Col1)
                                .Select(g =>
                                {
                                    var first = g.First();
                                    var bed = g.Sum(x => x.Col3);
                                    var bes = g.Sum(x => x.Col4);

                                    var Mande = bed - bes;
                                    if (Mande > 0)
                                    {
                                        bed = Mande;
                                        bes = 0;
                                    }
                                    if (Mande < 0)
                                    {
                                        bed = 0;
                                        bes = Math.Abs(Mande);
                                    }
                                    if (Mande == 0)
                                    {
                                        bed = 0;
                                        bes = 0;
                                    }

                                    return new ExcelRow
                                    {
                                        Col1 = first.Col1,
                                        Col2 = first.Col2,
                                        Col3 = bed,
                                        Col4 = bes
                                    };
                                }).ToList();
                return merged;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to MergeDuplicateRows");
                throw;
            }

        }

        public async Task<MemoryStream> GenerateRawTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, string FolderPath)
        {
            try
            {

                _logger.LogInformation("Starting GenerateRawTablesAsync...");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 2;

                foreach (var item in financialRecords)
                {
                    worksheet.Cell(row, 1).Value = item.Kol_Code;
                    worksheet.Cell(row, 2).Value = item.Kol_Title;
                    worksheet.Cell(row, 3).Value = item.Moeen_Code;
                    worksheet.Cell(row, 4).Value = item.Moeen_Title;
                    worksheet.Cell(row, 3).Value = item.Mande_Bed;
                    worksheet.Cell(row, 4).Value = item.Mande_Bes;
                    row++;
                }

                var stream = new MemoryStream();
                return await Task.FromResult(stream);
            }
            catch(Exception ex)
            {
                _logger.LogError(ex, "Failed to GenerateRawTablesAsync");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { "خطا در تولید جدول تراز خام" }
                    },
                FolderPath
                );
            }
        }

        public async Task<MemoryStream> GeneratePoyaTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, string FolderPath)
        {
            try
            {
                _logger.LogInformation("Starting GeneratePoyaTablesAsync...");

                var workbook = excelExporter.GetWorkbook();
                var stream = await GenerateRawTablesAsync(financialRecords, excelExporter, workbook, FolderPath);
                stream.Position = 0;


                var worksheet = workbook.Worksheets.Add("تراز اکسیر");
                worksheet.RightToLeft = true;
                int row = 2;

                var rows = financialRecords.Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}_{x.Moeen_Code}",
                    Col2 = $"{x.Kol_Title}_{x.Moeen_Title}",
                    Col3 = x.Mande_Bed,
                    Col4 = x.Mande_Bes,
                }).ToList();

                var mergedRows = MergeDuplicateRows(rows);

                var duplicateKeys = mergedRows
                                    .GroupBy(r => r.Col1)
                                    .Where(g => g.Count() > 1)
                                    .Select(g => g.Key)
                                    .ToList();

                if (duplicateKeys.Any())
                {
                    var dupList = string.Join(", ", duplicateKeys);
                    _logger.LogWarning($"Duplicate values found in Col1: {dupList}");
                    mergedRows = MergeDuplicateRows(mergedRows);
                }

                var emptyCol2 = mergedRows.Where(r => string.IsNullOrWhiteSpace(r.Col2)).ToList();

                if (emptyCol2.Any())
                {
                    _logger.LogWarning($"Found {emptyCol2.Count} rows with empty Col2.");
                    foreach (var item in emptyCol2)
                    {
                        item.Col2 = item.Col1;
                    }
                }

                decimal totalBed = mergedRows.Sum(r => r.Col3);
                decimal totalBes = mergedRows.Sum(r => r.Col4);

                if (totalBed != totalBes)
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"مانده بدهکار {totalBed} و مانده بستانکار {totalBes} برابر نمی باشد.  " }
                        },
                    FolderPath
                    );
                }


                foreach (var item in mergedRows)
                {
                    worksheet.Cell(row, 1).Value = item.Col1;
                    worksheet.Cell(row, 2).Value = item.Col2;
                    worksheet.Cell(row, 3).Value = item.Col3;
                    worksheet.Cell(row, 4).Value = item.Col4;
                    row++;
                }

                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch(Exception ex)
            {
                _logger.LogError(ex, "Failed to GeneratePoyaTablesAsync");
                throw;
            }
        }

        public async Task<MemoryStream> GenerateRayanTablesAsync(List<RayanFinancialRecord> RayanFinancialRecord, IExcelExporter excelExporter, string FolderPath)
        {
            try
            {
                _logger.LogInformation("Starting GenerateRayanTablesAsync...");

                var workbook = excelExporter.GetWorkbook();
                var stream = await GenerateRawRayanTablesAsync(RayanFinancialRecord, excelExporter, workbook, FolderPath);
                stream.Position = 0;


                var worksheet = workbook.Worksheets.Add("تراز اکسیر");
                worksheet.RightToLeft = true;
                int row = 2;

                var rows = RayanFinancialRecord.Select(x =>
                {

                    var code = $"{x.Kol_Code}_{x.Moeen_Code[^3..]}_{x.Tafsili_Code}";
                    var title = $"{x.Kol_Title}_{x.Moeen_Title}_{x.Tafsili_Title}";

                    if (x.joze1_Code.Length == 17)
                    { 
                        code += $"_{x.joze1_Code[^6..]}";
                        title += $"_{x.joze1_Title}";

                        if (x.joze1_Code.Length == 21)
                        {
                            code += $"_{x.joze1_Code[^4..]}";
                            title += $"_{x.joze1_Title}";
                        }
                    }

                    return new ExcelRow
                    {
                        Col1 = code,
                        Col2 = title,
                        Col3 = x.Mande_Bed,
                        Col4 = x.Mande_Bes,
                    };

                }).ToList();

                var mergedRows = MergeDuplicateRows(rows);

                var duplicateKeys = mergedRows
                                    .GroupBy(r => r.Col1)
                                    .Where(g => g.Count() > 1)
                                    .Select(g => g.Key)
                                    .ToList();

                if (duplicateKeys.Any())
                {
                    var dupList = string.Join(", ", duplicateKeys);
                    _logger.LogWarning($"Duplicate values found in Col1: {dupList}");
                    mergedRows = MergeDuplicateRows(mergedRows);
                }

                var emptyCol2 = mergedRows.Where(r => string.IsNullOrWhiteSpace(r.Col2)).ToList();

                if (emptyCol2.Any())
                {
                    _logger.LogWarning($"Found {emptyCol2.Count} rows with empty Col2.");
                    foreach (var item in emptyCol2)
                    {
                        item.Col2 = item.Col1;
                    }
                }

                decimal totalBed = mergedRows.Sum(r => r.Col3);
                decimal totalBes = mergedRows.Sum(r => r.Col4);

                if (totalBed != totalBes)
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"مانده بدهکار {totalBed} و مانده بستانکار {totalBes} برابر نمی باشد.  " }
                        },
                    FolderPath
                    );
                }


                foreach (var item in mergedRows)
                {
                    worksheet.Cell(row, 1).Value = item.Col1;
                    worksheet.Cell(row, 2).Value = item.Col2;
                    worksheet.Cell(row, 3).Value = item.Col3;
                    worksheet.Cell(row, 4).Value = item.Col4;
                    row++;
                }

                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch(Exception ex)
            {
                _logger.LogError(ex, "Failed to GenerateRayanTablesAsync");
                throw;
            }
        }
        public async Task<MemoryStream> GenerateRawRayanTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, string FolderPath)
        {
            try
            {

                _logger.LogInformation("Starting GenerateRawRAyanTablesAsync...");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 2;

                foreach (var item in financialRecords)
                {
                    worksheet.Cell(row, 1).Value = item.Kol_Code;
                    worksheet.Cell(row, 2).Value = item.Kol_Title;
                    worksheet.Cell(row, 3).Value = item.Moeen_Code;
                    worksheet.Cell(row, 4).Value = item.Moeen_Title;
                    worksheet.Cell(row, 5).Value = item.Tafsili_Code;
                    worksheet.Cell(row, 6).Value = item.Tafsili_Title;
                    worksheet.Cell(row, 7).Value = item.joze1_Code;
                    worksheet.Cell(row, 8).Value = item.joze1_Title;
                    worksheet.Cell(row, 9).Value = item.joze2_Code;
                    worksheet.Cell(row, 10).Value = item.joze2_Title;
                    worksheet.Cell(row, 11).Value = item.Mande_Bed;
                    worksheet.Cell(row, 12).Value = item.Mande_Bes;
                    row++;
                }

                var stream = new MemoryStream();
                return await Task.FromResult(stream);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to GenerateRawRayanTablesAsync");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { "خطا در تولید جدول تراز خام" }
                    },
                FolderPath
                );
            }
        }
    }
}

