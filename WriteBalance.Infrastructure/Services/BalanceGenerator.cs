using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Azure;
using Azure.Core;
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

        public async Task<MemoryStream> GenerateTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateTablesAsync"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();

                var streamReport = await GenerateRawTablesAsync(financialRecords, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                var rows = financialRecords.Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}_{x.Moeen_Code}",
                    Col2 = $"{x.Kol_Title}_{x.Moeen_Title}",
                    Col3 = x.Remain_First_Debit ?? decimal.Zero,
                    Col4 = x.Remain_First_Credit?? decimal.Zero,
                    Col5 = x.Flow_Debit ?? decimal.Zero,
                    Col6 = x.Flow_Credit ?? decimal.Zero,
                }).ToList();

                var rowsEditRemain = await Calculate_New_rows(rows);
                var mergedRows = MergeDuplicateRows(rowsEditRemain);

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
                var ekhtelaf = totalBed - totalBes;

                if (totalBed != totalBes)
                {
                    if (Math.Abs(ekhtelaf) > 100)
                    {
                        excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");
                        Logger.WriteEntry(JsonConvert.SerializeObject($"Not Balance with  {ekhtelaf}"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");

                        string formatted = ekhtelaf.ToString("#,##0.##");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"تراز به مقدار {formatted} بالانس نمیباشد." }
                            },
                        requestDB.FolderPath
                        );
                    }
                    else if (Math.Abs(ekhtelaf) <= 100)
                    {
                        if (totalBed > totalBes)
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = Math.Abs(ekhtelaf),
                                Col4 = 0,
                            });
                        }
                        else
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = 0,
                                Col4 = Math.Abs(ekhtelaf),
                            });
                        }
                    }

                }

                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Col3 - item.Col4 == 0)
                    {
                        continue;
                    }
                    else
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
                        writeValue++;
                    }

                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;


                var range = worksheetReport.Range("K:P");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheetReport.Range("A1:P1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;

                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch (ConnectionMessageException ex)
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
                                    if (Mande >= 0)
                                    {
                                        bed = Mande;
                                        bes = 0;
                                    }
                                    else if (Mande < 0)
                                    {
                                        bed = 0;
                                        bes = Math.Abs(Mande);
                                    }

                                    return new ExcelRow
                                    {
                                        Col1 = first.Col1,
                                        Col2 = first.Col2,
                                        Col3 = bed,
                                        Col4 = bes,
                                        Col5 = 0,
                                        Col6 = 0

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
        public async Task<MemoryStream> GenerateRawTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawTablesAsync"), $"BalanceGenerator:GenerateRawTablesAsync --typeReport:Info");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 1;

                worksheet.Cell(row, 1).Value = "Kol_Code";
                worksheet.Cell(row, 2).Value = "Kol_Title";
                worksheet.Cell(row, 3).Value = "Moeen_Code";
                worksheet.Cell(row, 4).Value = "Moeen_Title";
                worksheet.Cell(row, 5).Value = "Tafsil_Code";
                worksheet.Cell(row, 6).Value = "Tafsil_Tilte";
                worksheet.Cell(row, 7).Value = "FinApplication_Title";
                worksheet.Cell(row, 8).Value = "AccountNature_ID";
                worksheet.Cell(row, 9).Value = "AccountNature_Title";
                worksheet.Cell(row, 10).Value = "Motamam";
                worksheet.Cell(row, 11).Value = "Remain_First_Credit";
                worksheet.Cell(row, 12).Value = "Remain_First_Debit";
                worksheet.Cell(row, 13).Value = "Flow_Credit";
                worksheet.Cell(row, 14).Value = "Flow_Debit";
                worksheet.Cell(row, 15).Value = "Remain_Last_Credit";
                worksheet.Cell(row, 16).Value = "Remain_last_Debit";
                row = 2;
                int writeValue = 0;

                foreach (var item in financialRecords)
                {
                    if (requestDB.AllOrHasMandeh == "2" && await Calculate_Last_Remain(item))
                    {
                        continue;
                    }
                    else
                    {
                        worksheet.Cell(row, 1).Value = item.Kol_Code;
                        worksheet.Cell(row, 2).Value = item.Kol_Title;
                        worksheet.Cell(row, 3).Value = item.Moeen_Code;
                        worksheet.Cell(row, 4).Value = item.Moeen_Title;
                        worksheet.Cell(row, 5).Value = item.Tafzil_Code;
                        worksheet.Cell(row, 6).Value = item.Tafzil_Tilte;
                        worksheet.Cell(row, 7).Value = item.FinApplication_Title;
                        worksheet.Cell(row, 8).Value = item.AccountNature_ID;
                        worksheet.Cell(row, 9).Value = item.AccountNature_Title;
                        worksheet.Cell(row, 10).Value = item.Motamam;
                        worksheet.Cell(row, 11).Value = item.Remain_First_Credit;
                        worksheet.Cell(row, 12).Value = item.Remain_First_Debit;
                        worksheet.Cell(row, 13).Value = item.Flow_Credit;
                        worksheet.Cell(row, 14).Value = item.Flow_Debit;
                        worksheet.Cell(row, 15).Value = item.Remain_Last_Credit;
                        worksheet.Cell(row, 16).Value = item.Remain_last_Debit;

                        row++;

                        writeValue++;
                    }
                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                worksheet.Style.Font.FontName = "B Nazanin";
                worksheet.Style.Font.FontSize = 11;

                var range = worksheet.Range("K:P");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheet.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheet.RangeUsed();

                if (usedRange != null)
                {
                    worksheet.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheet.Range("A1:P1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;


                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:GenerateRawTablesAsync --typeReport:Error");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { "خطا در تولید جدول تراز خام" }
                    },
                requestDB.FolderPath
                );
            }
        }

        public async Task<List<ExcelRow>> Calculate_New_rows(List<ExcelRow> Rows)
        {
            try
            {
                foreach (ExcelRow row in Rows)
                {
                    decimal bed = 0;
                    decimal bes = 0;

                    // mandeh bedehkar
                    if (row.Col3 < 0)
                    {
                        bes += Math.Abs(row.Col3);
                    }
                    else if (row.Col3 >= 0)
                    {
                        bed += Math.Abs(row.Col3);
                    }

                    // mandeh bestankar
                    if (row.Col4 < 0)
                    {
                        bed += Math.Abs(row.Col4);
                    }
                    else if (row.Col4 >= 0)
                    {
                        bes += Math.Abs(row.Col4);
                    }

                    // gardesh bedehkar
                    if (row.Col5 < 0)
                    {
                        bes += Math.Abs(row.Col5 ?? decimal.Zero);
                    }
                    if (row.Col5 >= 0)
                    {
                        bed += Math.Abs(row.Col5 ?? decimal.Zero);
                    }

                    //gardesh bestankar
                    if (row.Col6 < 0)
                    {
                        bed += Math.Abs(row.Col6 ?? decimal.Zero);
                    }
                    else if (row.Col6 >= 0)
                    {
                        bes += Math.Abs(row.Col6 ?? decimal.Zero);
                    }

                    // mandeh
                    if (bed - bes >= 0)
                    {
                        row.Col3 = Math.Abs(bed - bes);
                        row.Col4 = 0;
                        row.Col5 = 0;
                        row.Col6 = 0;
                    }
                    else if(bed - bes < 0)
                    {
                        row.Col3 = 0;
                        row.Col4 = Math.Abs(bed - bes);
                        row.Col5 = 0;
                        row.Col6 = 0;
                    }
                }

                return await Task.FromResult(Rows);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:Calculate_New_rows --typeReport:Error");
                throw;
            }
        }
        public async Task<bool> Calculate_Last_Remain(FinancialRecord Record)
        {
            try
            {
                decimal bed = 0;
                decimal bes = 0;

                // mandeh bedehkar
                if (Record.Remain_First_Debit < 0)
                {
                    bes += Math.Abs(Record.Remain_First_Debit ?? decimal.Zero);
                }
                else if (Record.Remain_First_Debit >= 0)
                {
                    bed += Math.Abs(Record.Remain_First_Debit ?? decimal.Zero);
                }

                // mandeh bestankar
                if (Record.Remain_First_Credit < 0)
                {
                    bed += Math.Abs(Record.Remain_First_Credit ?? decimal.Zero);
                }
                else if (Record.Remain_First_Credit >= 0)
                {
                    bes += Math.Abs(Record.Remain_First_Credit ?? decimal.Zero);
                }

                // gardesh bedehkar
                if (Record.Flow_Debit < 0)
                {
                    bes += Math.Abs(Record.Flow_Debit ?? decimal.Zero);
                }
                if (Record.Flow_Debit >= 0)
                {
                    bed += Math.Abs(Record.Flow_Debit ?? decimal.Zero);
                }

                //gardesh bestankar
                if (Record.Flow_Credit < 0)
                {
                    bed += Math.Abs(Record.Flow_Credit ?? decimal.Zero);
                }
                else if (Record.Flow_Credit >= 0)
                {
                    bes += Math.Abs(Record.Flow_Credit ?? decimal.Zero);
                }

                // mandeh
                if (bed - bes == 0)
                {
                    return await Task.FromResult(true);
                }
                else
                {
                    return await Task.FromResult(false);
                }
            }
            catch (Exception ex) 
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:Calculate_Last_Remain --typeReport:Error");
                throw;
            }

        }
        /*
        public async Task<MemoryStream> GenerateTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateTablesAsync"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();

                var streamReport = await GenerateRawTablesAsync(financialRecords, excelExporter, workbookReport, requestDB);
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
                var ekhtelaf = totalBed - totalBes;

                if (totalBed != totalBes)
                {
                    if (Math.Abs(ekhtelaf) > 100)
                    {
                        excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");
                        Logger.WriteEntry(JsonConvert.SerializeObject($"Not Balance with  {ekhtelaf}"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");

                        string formatted = ekhtelaf.ToString("#,##0.##");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"تراز به مقدار {formatted} بالانس نمیباشد." }
                            },
                        requestDB.FolderPath
                        );
                    }
                    else if (Math.Abs(ekhtelaf) <= 100)
                    {
                        if (totalBed > totalBes)
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = Math.Abs(ekhtelaf),
                                Col4 = 0,
                            });
                        }
                        else
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = 0,
                                Col4 = Math.Abs(ekhtelaf),
                            });
                        }
                    }

                }

                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Col3 - item.Col4 == 0)
                    {
                        continue;
                    }
                    else
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
                        writeValue++;
                    }

                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;


                var range = worksheetReport.Range("H:K");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheetReport.Range("A1:k1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;

                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch (ConnectionMessageException ex)
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
        public async Task<MemoryStream> GenerateRawTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, DBRequestDto requestDB)
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
                worksheet.Cell(row, 5).Value = "کد تفضیلی";
                worksheet.Cell(row, 6).Value = "عنوان تفضیلی";
                worksheet.Cell(row, 7).Value = "عنوان تراز";
                worksheet.Cell(row, 8).Value = "گردش بدهکار";
                worksheet.Cell(row, 9).Value = "گردش بستانکار";
                worksheet.Cell(row, 10).Value = "مانده بدهکار";
                worksheet.Cell(row, 11).Value = "مانده بستانکار";
                row = 2;
                int writeValue = 0;

                foreach (var item in financialRecords)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Mande_Bed - item.Mande_Bes == 0) 
                    {
                        continue;
                    }
                    else
                    {
                        worksheet.Cell(row, 1).Value = item.Kol_Code;
                        worksheet.Cell(row, 2).Value = item.Kol_Title;
                        worksheet.Cell(row, 3).Value = item.Moeen_Code;
                        worksheet.Cell(row, 4).Value = item.Moeen_Title;
                        worksheet.Cell(row, 5).Value = item.Tafzil_Code;
                        worksheet.Cell(row, 6).Value = item.Tafzil_Tilte;
                        worksheet.Cell(row, 7).Value = item.FinApplication_Title;
                        worksheet.Cell(row, 8).Value = item.Gardersh_Bed;
                        worksheet.Cell(row, 9).Value = item.Gardersh_Bes;
                        worksheet.Cell(row, 10).Value = item.Mande_Bed;
                        worksheet.Cell(row, 11).Value = item.Mande_Bes;
                        row++;

                        writeValue++;
                    }
                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                worksheet.Style.Font.FontName = "B Nazanin";
                worksheet.Style.Font.FontSize = 11;

                var range = worksheet.Range("H:K");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheet.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheet.RangeUsed();

                if (usedRange != null)
                {
                    worksheet.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheet.Range("A1:k1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;


                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:GenerateRawTablesAsync --typeReport:Error");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { "خطا در تولید جدول تراز خام" }
                    },
                requestDB.FolderPath
                );
            }
        }
        */
        public async Task GeneratePoyaTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GeneratePoyaTablesAsync"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();
                var workbookUploadArzi = excelExporter.GetWorkbookUploadArzi();
                var streamReport = await GenerateRawPouyaTablesAsync(financialRecords, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                var rowsRial = financialRecords.Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}_{x.Arz_Code}_{x.Moeen_Code}",
                    Col2 = $"{x.Kol_Title}_{x.Sharh_Arz}",
                    Col3 = x.Mande_Bed_rial?? 0,
                    Col4 = x.Mande_Bes_rial ?? 0,
                }).ToList();

                var mergedRows = MergeDuplicateRows(rowsRial);
                mergedRows = await checkBalance(mergedRows, excelExporter, requestDB, streamReport);

                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر ریالی");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Col3 - item.Col4 == 0)
                    {
                        continue;
                    }
                    else
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
                        writeValue++;
                    }
                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }


                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;

                var range = worksheetReport.Range("L:O");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("K").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheetReport.Range("A1:O1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;
                streamReport.Position = 0;

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                await excelExporter.SaveUploadAsync(streamUpload, requestDB.FolderPath, requestDB.FileNameRial);

                var rowsArzi = financialRecords.Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}_{x.Arz_Code}_{x.Moeen_Code}",
                    Col2 = $"{x.Kol_Title}_{x.Sharh_Arz}",
                    Col3 = x.Mande_Bed_arzi ?? 0,
                    Col4 = x.Mande_Bes_arzi ?? 0,
                }).ToList();


                var worksheetUploadArzi = workbookUploadArzi.Worksheets.Add("Data");
                var worksheetReportArzi = workbookReport.Worksheets.Add("تراز اکسیر ارزی");
                worksheetUploadArzi.RightToLeft = true;
                worksheetReportArzi.RightToLeft = true;
                row = 2;
                writeValue = 0;

                foreach (var item in mergedRows)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Col3 - item.Col4 == 0)
                    {
                        continue;
                    }
                    else
                    {
                        worksheetUploadArzi.Cell(row, 1).Value = item.Col1;
                        worksheetUploadArzi.Cell(row, 2).Value = item.Col2;
                        worksheetUploadArzi.Cell(row, 3).Value = item.Col3.ToString();
                        worksheetUploadArzi.Cell(row, 4).Value = item.Col4.ToString();

                        worksheetReportArzi.Cell(row, 1).Value = item.Col1;
                        worksheetReportArzi.Cell(row, 2).Value = item.Col2;
                        worksheetReportArzi.Cell(row, 3).Value = item.Col3;
                        worksheetReportArzi.Cell(row, 4).Value = item.Col4; ;

                        row++;
                        writeValue++;
                    }
                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }


                worksheetReportArzi.Style.Font.FontName = "B Nazanin";
                worksheetReportArzi.Style.Font.FontSize = 11;

                range = worksheetReportArzi.Range("L:O");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReportArzi.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReportArzi.Column("C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReportArzi.Column("K").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                usedRange = worksheetReportArzi.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReportArzi.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                headerRange = worksheetReportArzi.Range("A1:O1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;

                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUploadArzi = new MemoryStream();
                workbookUploadArzi.SaveAs(streamUploadArzi);
                streamUploadArzi.Position = 0;
                await excelExporter.SaveUploadArziAsync(streamUploadArzi, requestDB.FolderPath, requestDB.FileNameArzi);

            }
            catch (ConnectionMessageException ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("GeneratePoyaTablesAsync failed!"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");
                throw;
            }
        }
        public async Task<List<ExcelRow>> checkBalance(List<ExcelRow> mergedRows, IExcelExporter excelExporter, DBRequestDto requestDB, MemoryStream streamReport)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("checkBalance Starting!"), $"BalanceGenerator:checkBalance --typeReport:Info");
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
                var ekhtelaf = totalBed - totalBes;

                if (totalBed != totalBes)
                {
                    if (Math.Abs(ekhtelaf) > 100)
                    {
                        excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");
                        Logger.WriteEntry(JsonConvert.SerializeObject($"Not Balance with  {ekhtelaf}"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                        string formatted = ekhtelaf.ToString("#,##0.##");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"تراز به مقدار {formatted} بالانس نمیباشد." }
                            },
                        requestDB.FolderPath
                        );
                    }
                    else if (Math.Abs(ekhtelaf) <= 100)
                    {
                        if (totalBed > totalBes)
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = Math.Abs(ekhtelaf),
                                Col4 = 0,
                            });
                        }
                        else
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = 0,
                                Col4 = Math.Abs(ekhtelaf),
                            });
                        }
                    }
                }
                return await Task.FromResult(mergedRows); 
            }
            catch (ConnectionMessageException ex) {
                throw;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:checkBalance --typeReport:Error");
                throw;
            }
        }
        public async Task<MemoryStream> GenerateRawPouyaTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawPouyaTablesAsync"), $"BalanceGenerator:GenerateRawPouyaTablesAsync --typeReport:Info");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 1;

                worksheet.Cell(row, 1).Value = "تاریخ انتهای بازه گزارش گیری";
                worksheet.Cell(row, 2).Value = "کد شعبه";
                worksheet.Cell(row, 3).Value = "کد کل از دید بانک مرکزی ";
                worksheet.Cell(row, 4).Value = "عنوان کد کل";
                worksheet.Cell(row, 5).Value = "کد حساب";
                worksheet.Cell(row, 6).Value = "سرفصل کل";
                worksheet.Cell(row, 7).Value = "کد ارز";
                worksheet.Cell(row, 8).Value = "گروه معین";
                worksheet.Cell(row, 9).Value = "معین";
                worksheet.Cell(row, 10).Value = "تفصیلی";
                worksheet.Cell(row, 11).Value = "کد اختصاری ارز";
                worksheet.Cell(row, 12).Value = "شرح ارز";
                worksheet.Cell(row, 13).Value = "مانده بدهکار ارزی";
                worksheet.Cell(row, 14).Value = "مانده بستانکار ارزی";
                worksheet.Cell(row, 15).Value = "مانده بدهکار ریالی";
                worksheet.Cell(row, 16).Value = "مانده بستانکار ریالی";
                worksheet.Cell(row, 17).Value = "گردش بدهکار ریالی";
                worksheet.Cell(row, 18).Value = "گردش بستانکار ریالی";
                worksheet.Cell(row, 19).Value = "گردش بدهکاری ارزی";
                worksheet.Cell(row, 20).Value = "گردش بستانکار ارزی";

                row = 2;
                int writeValue = 0;

                foreach (var item in financialRecords)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Mande_Bed_arzi - item.Mande_Bes_arzi == 0)
                    {
                        continue;
                    }
                    else
                    {
                        worksheet.Cell(row, 1).Value = item.Taraz_Date;
                        worksheet.Cell(row, 2).Value = 0;
                        worksheet.Cell(row, 3).Value = item.Kol_Code_Markazi;
                        worksheet.Cell(row, 4).Value = item.Kol_Title;
                        worksheet.Cell(row, 5).Value = item.Hesab_Code;
                        worksheet.Cell(row, 6).Value = item.Kol_Code;
                        worksheet.Cell(row, 7).Value = item.Arz_Code;
                        worksheet.Cell(row, 8).Value = item.Moeen_Code;
                        worksheet.Cell(row, 9).Value = item.Moeen;
                        worksheet.Cell(row, 10).Value = item.Tafzili;
                        worksheet.Cell(row, 11).Value = item.Code_Arz_Abbr;
                        worksheet.Cell(row, 12).Value = item.Sharh_Arz;
                        worksheet.Cell(row, 13).Value = item.Mande_Bed_arzi;
                        worksheet.Cell(row, 14).Value = item.Mande_Bes_arzi;
                        worksheet.Cell(row, 15).Value = item.Mande_Bed_rial;
                        worksheet.Cell(row, 16).Value = item.Mande_Bes_rial;
                        worksheet.Cell(row, 17).Value = item.Gardersh_Bed_rial;
                        worksheet.Cell(row, 18).Value = item.Gardersh_Bes_rial;
                        worksheet.Cell(row, 19).Value = item.Gardersh_Bed_arzi;
                        worksheet.Cell(row, 20).Value = item.Gardersh_Bes_arzi;

                        row++;

                        writeValue++;
                    }
                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GenerateRawPouyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                worksheet.Style.Font.FontName = "B Nazanin";
                worksheet.Style.Font.FontSize = 11;

                var range = worksheet.Range("M:T");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheet.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("L").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheet.RangeUsed();

                if (usedRange != null)
                {
                    worksheet.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheet.Range("A1:T1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;

                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:GenerateRawPouyaTablesAsync --typeReport:Error");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { "خطا در تولید تراز خام" }
                    },
                requestDB.FolderPath
                );
            }
        }
        public async Task<MemoryStream> GenerateRayanTablesAsync(List<RayanFinancialRecord> RayanFinancialRecord, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRayanTablesAsync"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();
                var streamReport = await GenerateRawRayanTablesAsync(RayanFinancialRecord, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                var rows = RayanFinancialRecord.Select(x =>
                {

                    var code = $"{x.Kol_Code}_{x.Moeen_Code[^3..]}_{x.Tafsili_Code[^4..]}";
                    var title = $"{x.Kol_Title}_{x.Moeen_Title}_{x.Tafsili_Title}";

                    if (x.joze1_Code.Length == 17)
                    {
                        code += $"_{x.joze1_Code[^6..]}";
                        title += $"_{x.joze1_Title}";


                    }
                    else
                    {
                        code += $"_0";
                    }

                    if (x.joze2_Code.Length == 21)
                    {
                        code += $"_{x.joze2_Code[^4..]}";
                        title += $"_{x.joze2_Title}";
                    }
                    else
                    {
                        code += $"_0";
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
                var ekhtelaf = totalBed - totalBes;

                if (totalBed != totalBes)
                {
                    if (Math.Abs(ekhtelaf) > 100)
                    {
                        excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");
                        Logger.WriteEntry(JsonConvert.SerializeObject($"Not Balance with  {ekhtelaf}"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Error");

                        string formatted = ekhtelaf.ToString("#,##0.##");

                        throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"تراز به مقدار {formatted} بالانس نمیباشد." }
                            },
                        requestDB.FolderPath
                        );
                    }
                    else if (Math.Abs(ekhtelaf) <= 100)
                    {
                        if (totalBed > totalBes)
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = 0,
                                Col4 = Math.Abs(ekhtelaf),
                            });
                        }
                        else
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "بالانس",
                                Col3 = Math.Abs(ekhtelaf),
                                Col4 = 0,
                            });
                        }
                    }

                }

                var worksheetUpload = workbookUpload.Worksheets.Add("data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Col3 - item.Col4 == 0)
                    {
                        continue;
                    }
                    else
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
                        writeValue++;
                    }
                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;

                var range = worksheetReport.Range("R:V");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("H").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("J").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("L").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("O").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheetReport.Column("Q").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
                var headerRange = worksheetReport.Range("A1:V1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;


                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch (ConnectionMessageException ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Failed to GenerateRayanTablesAsync"), $"BalanceGenerator:GenerateRayanTablesAsync --typeReport:Error");
                throw;
            }
        }
        public async Task<MemoryStream> GenerateRawRayanTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawRayanTablesAsync"), $"BalanceGenerator:GenerateRawRayanTablesAsync --typeReport:Info");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 1;

                worksheet.Cell(row, 1).Value = "کد گروه";
                worksheet.Cell(row, 2).Value = "نام گروه";
                worksheet.Cell(row, 3).Value = "کد حساب کل";
                worksheet.Cell(row, 4).Value = "عنوان حساب کل";
                worksheet.Cell(row, 5).Value = "کد حساب معین";
                worksheet.Cell(row, 6).Value = "عنوان حساب معین";
                worksheet.Cell(row, 7).Value = "کد حساب تفصیلی";
                worksheet.Cell(row, 8).Value = "عنوان حساب تفصیلی";
                worksheet.Cell(row, 9).Value = "کد جز 1";
                worksheet.Cell(row, 10).Value = "عنوان جز 1";
                worksheet.Cell(row, 11).Value = "کد جز 2";
                worksheet.Cell(row, 12).Value = "عنوان جز 2";
                worksheet.Cell(row, 13).Value = "کد مرکز هزینه";
                worksheet.Cell(row, 14).Value = "کد واحد عملیاتی";
                worksheet.Cell(row, 15).Value = "نام واحد عملیاتی";
                worksheet.Cell(row, 16).Value = "کد پرونده";
                worksheet.Cell(row, 17).Value = "نام پرونده";
                worksheet.Cell(row, 18).Value = "مانده اول دوره";
                worksheet.Cell(row, 19).Value = "بدهکار";
                worksheet.Cell(row, 20).Value = "بستانکار";
                worksheet.Cell(row, 21).Value = "مانده بدهکار";
                worksheet.Cell(row, 22).Value = "مانده بستانکار";

                row = 2;
                int writeValue = 0;

                foreach (var item in financialRecords)
                {
                    if (requestDB.AllOrHasMandeh == "2" && item.Mande_Bed - item.Mande_Bes == 0)
                    {
                        continue;
                    }
                    else
                    {
                        worksheet.Cell(row, 1).Value = item.Group_code;
                        worksheet.Cell(row, 2).Value = item.Group_Title;
                        worksheet.Cell(row, 3).Value = item.Kol_Code;
                        worksheet.Cell(row, 4).Value = item.Kol_Title;
                        worksheet.Cell(row, 5).Value = item.Moeen_Code;
                        worksheet.Cell(row, 6).Value = item.Moeen_Title;
                        worksheet.Cell(row, 7).Value = item.Tafsili_Code;
                        worksheet.Cell(row, 8).Value = item.Tafsili_Title;
                        worksheet.Cell(row, 9).Value = item.joze1_Code;
                        worksheet.Cell(row, 10).Value = item.joze1_Title;
                        worksheet.Cell(row, 11).Value = item.joze2_Code;
                        worksheet.Cell(row, 12).Value = item.joze2_Title;
                        worksheet.Cell(row, 13).Value = item.Code_Markaz_Hazineh;
                        worksheet.Cell(row, 14).Value = item.Code_Vahed_Amaliyat;
                        worksheet.Cell(row, 15).Value = item.Name_Vahed_Amaliyat;
                        worksheet.Cell(row, 16).Value = item.Code_Parvandeh;
                        worksheet.Cell(row, 17).Value = item.Name_Parvandeh;
                        worksheet.Cell(row, 18).Value = item.Mandeh_Aval_dore;
                        worksheet.Cell(row, 19).Value = item.bedehkar;
                        worksheet.Cell(row, 20).Value = item.bestankar;
                        worksheet.Cell(row, 21).Value = item.Mande_Bed;
                        worksheet.Cell(row, 22).Value = item.Mande_Bes;
                        row++;

                        writeValue++;
                    }
                }

                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                worksheet.Style.Font.FontName = "B Nazanin";
                worksheet.Style.Font.FontSize = 11;

                var range = worksheet.Range("R:V");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheet.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("D").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("F").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("H").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("J").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("L").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("O").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("Q").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheet.RangeUsed();

                if (usedRange != null)
                {
                    worksheet.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheet.Range("A1:V1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;


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
                requestDB.FolderPath
                );
            }
        }

    }
}

