using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Infrastructure.Services
{
    public class BalanceGenerator : IBalanceGenerator
    {
        public async Task<MemoryStream> GenerateTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, string FolderPath)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateTablesAsync"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload(); 
                var streamReport = await GenerateRawTablesAsync(financialRecords, excelExporter, workbookReport, FolderPath);
                streamReport.Position = 0;

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
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Duplicate values found in Col1: {dupList}"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Warning");
                    mergedRows = MergeDuplicateRows(mergedRows);
                }

                var emptyCol2 = mergedRows.Where(r => string.IsNullOrWhiteSpace(r.Col2)).ToList();

                if (emptyCol2.Any())
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Found {emptyCol2.Count} rows with empty Col2."), $"BalanceGenerator:GenerateTablesAsync --typeReport:Warning");
                    foreach (var item in emptyCol2)
                    {
                        item.Col2 = item.Col1;
                    }
                }

                decimal totalBed = mergedRows.Sum(r => r.Col3);
                decimal totalBes = mergedRows.Sum(r => r.Col4);

                if (totalBed != totalBes)
                {
                    excelExporter.SaveReportAsync(streamReport, FolderPath, "Raw_Balance.xlsx");
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Found {emptyCol2.Count} rows with empty Col2."), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");

                    var ekhtelaf = Math.Abs(totalBed - totalBes);
                    string formatted = ekhtelaf.ToString("#,##0.##");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تراز به مقدار {formatted} بالانس نمیباشد." }
                        },
                    FolderPath
                    );
                }

                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;

                Logger.WriteEntry(JsonConvert.SerializeObject($"merged rows count:{mergedRows.Count}"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Info");

                foreach (var item in mergedRows)
                {
                    worksheetUpload.Cell(row, 1).Value = item.Col1;
                    worksheetUpload.Cell(row, 2).Value = item.Col2;
                    worksheetUpload.Cell(row, 3).Value = item.Col3.ToString();
                    worksheetUpload.Cell(row, 4).Value = item.Col4.ToString();

                    worksheetReport.Cell(row, 1).Value = item.Col1;
                    worksheetReport.Cell(row, 2).Value = item.Col2;
                    worksheetReport.Cell(row, 3).Value = item.Col3;
                    worksheetReport.Cell(row, 4).Value = item.Col4; ;

                    row++;
                }

                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                excelExporter.SaveReportAsync(streamReport, FolderPath, "Balance.xlsx");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync failed!"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");
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
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:MergeDuplicateRows --typeReport:Error");
                throw;
            }

        }

        public async Task<MemoryStream> GenerateRawTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, string FolderPath)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawTablesAsync"), $"BalanceGenerator:GenerateRawTablesAsync --typeReport:Info");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 1;


                worksheet.Cell(row, 1).Value = "کد حساب کل";
                worksheet.Cell(row, 2).Value = "عنوان حساب کل";
                worksheet.Cell(row, 3).Value = "کد حساب معین";
                worksheet.Cell(row, 4).Value = "عنوان حساب معین";
                worksheet.Cell(row, 5).Value = "بدهکار";
                worksheet.Cell(row, 6).Value = "بستانکار";
                row = 2;

                foreach (var item in financialRecords)
                {
                    worksheet.Cell(row, 1).Value = item.Kol_Code;
                    worksheet.Cell(row, 2).Value = item.Kol_Title;
                    worksheet.Cell(row, 3).Value = item.Moeen_Code;
                    worksheet.Cell(row, 4).Value = item.Moeen_Title;
                    worksheet.Cell(row, 5).Value = item.Mande_Bed;
                    worksheet.Cell(row, 6).Value = item.Mande_Bes;
                    row++;
                }

                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch(Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:GenerateRawTablesAsync --typeReport:Error");

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
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GeneratePoyaTablesAsync"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();
                var streamReport = await GenerateRawTablesAsync(financialRecords, excelExporter, workbookReport, FolderPath);
                streamReport.Position = 0;

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
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Duplicate values found in Col1: {dupList}"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Warning");
                    mergedRows = MergeDuplicateRows(mergedRows);
                }

                var emptyCol2 = mergedRows.Where(r => string.IsNullOrWhiteSpace(r.Col2)).ToList();

                if (emptyCol2.Any())
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Found {emptyCol2.Count} rows with empty Col2."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Warning");
                    foreach (var item in emptyCol2)
                    {
                        item.Col2 = item.Col1;
                    }
                }

                decimal totalBed = mergedRows.Sum(r => r.Col3);
                decimal totalBes = mergedRows.Sum(r => r.Col4);

                if (totalBed != totalBes)
                {
                    excelExporter.SaveReportAsync(streamReport, FolderPath, "Raw_Balance.xlsx");
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Found {emptyCol2.Count} rows with empty Col2."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    var ekhtelaf = Math.Abs(totalBed - totalBes);
                    string formatted = ekhtelaf.ToString("#,##0.##");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تراز به مقدار {formatted} بالانس نمیباشد." }
                        },
                    FolderPath
                    );
                }

                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;

                foreach (var item in mergedRows)
                {
                    worksheetUpload.Cell(row, 1).Value = item.Col1;
                    worksheetUpload.Cell(row, 2).Value = item.Col2;
                    worksheetUpload.Cell(row, 3).Value = item.Col3.ToString();
                    worksheetUpload.Cell(row, 4).Value = item.Col4.ToString();

                    worksheetReport.Cell(row, 1).Value = item.Col1;
                    worksheetReport.Cell(row, 2).Value = item.Col2;
                    worksheetReport.Cell(row, 3).Value = item.Col3;
                    worksheetReport.Cell(row, 4).Value = item.Col4; ;

                    row++;
                }

                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                excelExporter.SaveReportAsync(streamReport, FolderPath, "Balance.xlsx");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("GeneratePoyaTablesAsync failed!"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");
                throw;
            }
        }

        public async Task<MemoryStream> GenerateRayanTablesAsync(List<RayanFinancialRecord> RayanFinancialRecord, IExcelExporter excelExporter, string FolderPath)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRayanTablesAsync"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload(); 
                var streamReport = await GenerateRawRayanTablesAsync(RayanFinancialRecord, excelExporter, workbookReport, FolderPath);
                streamReport.Position = 0;

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
                        Col3 = (decimal)x.Mande_Bed,
                        Col4 = (decimal)x.Mande_Bes,
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
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Duplicate values found in Col1: {dupList}"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Warning");
                    mergedRows = MergeDuplicateRows(mergedRows);
                }

                var emptyCol2 = mergedRows.Where(r => string.IsNullOrWhiteSpace(r.Col2)).ToList();

                if (emptyCol2.Any())
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Found {emptyCol2.Count} rows with empty Col2."), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Warning");
                    foreach (var item in emptyCol2)
                    {
                        item.Col2 = item.Col1;
                    }
                }

                decimal totalBed = mergedRows.Sum(r => r.Col3);
                decimal totalBes = mergedRows.Sum(r => r.Col4);

                if (totalBed != totalBes)
                {
                    excelExporter.SaveReportAsync(streamReport, FolderPath, "Raw_Balance.xlsx");
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Found {emptyCol2.Count} rows with empty Col2."), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Error");

                    var ekhtelaf = Math.Abs(totalBed - totalBes);
                    string formatted = ekhtelaf.ToString("#,##0.##");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تراز به مقدار {formatted} بالانس نمیباشد." }
                        },
                    FolderPath
                    );
                }

                var worksheetUpload = workbookUpload.Worksheets.Add("data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;

                Logger.WriteEntry(JsonConvert.SerializeObject($"merged rows count:{mergedRows.Count}"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Info");


                foreach (var item in mergedRows)
                {
                    worksheetUpload.Cell(row, 1).Value = item.Col1;
                    worksheetUpload.Cell(row, 2).Value = item.Col2;
                    worksheetUpload.Cell(row, 3).Value = item.Col3.ToString();
                    worksheetUpload.Cell(row, 4).Value = item.Col4.ToString();

                    worksheetReport.Cell(row, 1).Value = item.Col1;
                    worksheetReport.Cell(row, 2).Value = item.Col2;
                    worksheetReport.Cell(row, 3).Value = item.Col3;
                    worksheetReport.Cell(row, 4).Value = item.Col4;

                    row++;
                }

                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                excelExporter.SaveReportAsync(streamReport, FolderPath, "Balance.xlsx");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Failed to GenerateRayanTablesAsync"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Error");
                throw;
            }
        }
        public async Task<MemoryStream> GenerateRawRayanTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, string FolderPath)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawRayanTablesAsync"), $"BalanceGenerator:GenerateRawRayanTablesAsync --typeReport:Info");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 1;

                worksheet.Cell(row, 1).Value = "کد حساب کل";
                worksheet.Cell(row, 2).Value = "عنوان حساب کل";
                worksheet.Cell(row, 3).Value = "کد حساب معین";
                worksheet.Cell(row, 4).Value = "عنوان حساب معین";
                worksheet.Cell(row, 5).Value = "کد حساب تفصیلی";
                worksheet.Cell(row, 6).Value = "عنوان حساب تفصیلی";
                worksheet.Cell(row, 7).Value = "کد جز 1";
                worksheet.Cell(row, 8).Value = "عنوان جز 1";
                worksheet.Cell(row, 9).Value = "کد جز 2";
                worksheet.Cell(row, 10).Value = "عنوان جز 2";
                worksheet.Cell(row, 11).Value = "بدهکار";
                worksheet.Cell(row, 12).Value = "بستانکار";

                row = 2;
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
                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:GenerateRawRayanTablesAsync --typeReport:Error");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { "خطا در تولید تراز خام" }
                    },
                FolderPath
                );
            }
        }
    }
}

