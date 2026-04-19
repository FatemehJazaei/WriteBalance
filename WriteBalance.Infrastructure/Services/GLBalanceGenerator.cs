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
    public class GLBalanceGenerator : IGLBalanceGenerator
    {

        public BalanceMerge _balanceMerge;
        public BalanceCheck _balanceCheck;
        public CalculateNewRows _calculateNewRows;
        public GLBalanceGenerator(BalanceMerge balanceMerge, BalanceCheck balanceCheck, CalculateNewRows calculateNewRows)
        {
            _balanceMerge = balanceMerge;
            _balanceCheck = balanceCheck;
            _calculateNewRows = calculateNewRows;
        }


        // تولید جدول تراز مانده برای همراه، سما و کاربردی 
        public async Task<MemoryStream> GenerateGLTablesAsync(List<GLFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateTablesAsync"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Info");
                // دو جدول اماده میشود، یکی برای گزارش و یکی برای اپلود 
                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();
                // تولید شیت خام برای گزارش دهی 
                var streamReport = await GenerateRawLGTablesAsync(financialRecords, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                List<ExcelRow> rows = new List<ExcelRow>();
                if (requestDB.TarazKolOrTarazMoeen == "1")
                {
                    // فیلتر کردن کد کل 6
                    rows = financialRecords
                        //.Where(x => (x.RBank_Code[0] != '6'))
                        .Select(x => new ExcelRow
                        {
                            Col1 = $"{x.RBank_Code[..4]}",
                            Col2 = $"{x.RBank_Title.Replace("***","_")}",
                            Col3 = x.Remain_last_Debit ?? decimal.Zero,
                            Col4 = x.Remain_Last_Credit ?? decimal.Zero,
                            Col5 = x.Flow_Debit ?? decimal.Zero,
                            Col6 = x.Flow_Credit ?? decimal.Zero,
                        }).ToList();

                }
                else if (requestDB.TarazKolOrTarazMoeen == "2" || requestDB.TarazKolOrTarazMoeen == "3")
                {
                    rows = financialRecords
                        //.Where(x => (x.RBank_Code[0] != '6'))
                        .Select(x => new ExcelRow
                        {
                            Col1 = $"{x.RBank_Code.Replace("-", "_")}",
                            Col2 = $"{x.RBank_Title.Replace("***", "_")}",
                            Col3 = x.Remain_last_Debit ?? decimal.Zero,
                            Col4 = x.Remain_Last_Credit ?? decimal.Zero,
                            Col5 = x.Flow_Debit ?? decimal.Zero,
                            Col6 = x.Flow_Credit ?? decimal.Zero,
                        }).ToList();
                }

                // محاسبه مانده برای هر رکورد
                var rowsEditRemain = await _calculateNewRows.Calculate_New_rows(rows);
                // یونیک کردن رکوردها بر اساس کد
                var mergedRows = _balanceMerge.MergeDuplicateRows(rowsEditRemain);

                // بررسی بالانس بودن  تراز 
                mergedRows = await _balanceCheck.checkBalance(mergedRows, excelExporter, requestDB, streamReport);

                // افوزدن شیت تراز محاسبه شده به  فایل اکسل اپلود  و فایل گزارش
                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                //سطر اول تراز خالی است 
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {
                    // بررسی گزینه انتخاب شده : همه رکورد ها یا فقط مانده دار ها 
                    // AllOrHasMandeh
                    // همه 1
                    // فقط مانده داره ها 2
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

                // در صورتی که همه رکورد ها بدون مانده باشند ، هیچ رکوردی در اکسل نوشته نمی شود
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

                //استایل دهی به گزارش 
                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;

                var range = worksheetReport.Range("C:D");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
                streamReport.Position = 0;


                // ذخیره فایل گزارش دهی 
                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch (ConnectionMessageException)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync failed!"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");
                throw;
            }

        }


        // تولید جدول برای تراز گردش همراه، سما و کاربردی 
        public async Task<MemoryStream> GenerateGardeshGLTablesAsync(List<GLFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateTablesAsync"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Info");
                // دو جدول اماده میشود، یکی برای گزارش و یکی برای اپلود 
                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();
                // تولید شیت خام برای گزارش دهی 
                var streamReport = await GenerateRawLGTablesAsync(financialRecords, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;


                List<ExcelRow> rows = new List<ExcelRow>();
                if (requestDB.TarazKolOrTarazMoeen == "1")
                {
                    // فیلتر کردن کد کل 6
                    rows = financialRecords
                        //.Where(x => (x.RBank_Code[0] != '6'))
                        .Select(x => new ExcelRow
                        {
                            Col1 = $"{x.RBank_Code[..4]}",
                            Col2 = $"{x.RBank_Title.Replace("***", "_")}",
                            Col3 = x.Remain_last_Debit ?? decimal.Zero,
                            Col4 = x.Remain_Last_Credit ?? decimal.Zero,
                            Col5 = x.Flow_Debit ?? decimal.Zero,
                            Col6 = x.Flow_Credit ?? decimal.Zero,
                        }).ToList();

                }
                else if (requestDB.TarazKolOrTarazMoeen == "2" || requestDB.TarazKolOrTarazMoeen == "3")
                {
                    rows = financialRecords
                        //.Where(x => (x.RBank_Code[0] != '6'))
                        .Select(x => new ExcelRow
                        {
                            Col1 = $"{x.RBank_Code.Replace("-", "_")}",
                            Col2 = $"{x.RBank_Title.Replace("***", "_")}",
                            Col3 = x.Remain_last_Debit ?? decimal.Zero,
                            Col4 = x.Remain_Last_Credit ?? decimal.Zero,
                            Col5 = x.Flow_Debit ?? decimal.Zero,
                            Col6 = x.Flow_Credit ?? decimal.Zero,
                        }).ToList();
                }

                // یونیک کردن رکوردها بر اساس کد
                var mergedRows = _balanceMerge.MergeDuplicateGardeshRows(rows);

                // بررسی بالانس بودن  تراز 
                mergedRows = await _balanceCheck.checkGardeshBalance(mergedRows, excelExporter, requestDB, streamReport);


                // افوزدن شیت تراز محاسبه شده به  فایل اکسل اپلود  و فایل گزارش
                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                //سطر اول تراز خالی است 
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {

                    worksheetUpload.Cell(row, 1).Value = item.Col1;
                    worksheetUpload.Cell(row, 2).Value = item.Col2;
                    worksheetUpload.Cell(row, 3).Value = item.Col5.ToString();
                    worksheetUpload.Cell(row, 4).Value = item.Col6.ToString();

                    worksheetReport.Cell(row, 1).Value = item.Col1;
                    worksheetReport.Cell(row, 2).Value = item.Col2;
                    worksheetReport.Cell(row, 3).Value = item.Col5;
                    worksheetReport.Cell(row, 4).Value = item.Col6;

                    row++;
                    writeValue++;

                }

                // در صورتی که همه رکورد ها بدون مانده باشند ، هیچ رکوردی در اکسل نوشته نمی شود
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

                //استایل دهی به گزارش 
                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;

                var range = worksheetReport.Range("C:D");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
                streamReport.Position = 0;

                // ذخیره فایل گزارش دهی 
                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                return await Task.FromResult(streamUpload);
            }
            catch (ConnectionMessageException ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync failed! {ex}"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");
                throw;
            }

        }

        // تولید جدول  تراز خام برای GL
        public async Task<MemoryStream> GenerateRawLGTablesAsync(List<GLFinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawLGTablesAsync"), $"BalanceGenerator:GenerateRawLGTablesAsync --typeReport:Info");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;
                int row = 1;

                // عنوان ستون ها  تنظیم میشود
                worksheet.Cell(row, 1).Value = "Branch_ID";
                worksheet.Cell(row, 2).Value = "RBank_Code";
                worksheet.Cell(row, 3).Value = "RBank_Title";
                worksheet.Cell(row, 4).Value = "FinApplication_ID";
                worksheet.Cell(row, 5).Value = "FinApplication_Title";
                worksheet.Cell(row, 6).Value = "Motamam";
                worksheet.Cell(row, 7).Value = "Remain_First_Credit";
                worksheet.Cell(row, 8).Value = "Remain_First_Debit";
                worksheet.Cell(row, 9).Value = "Flow_Credit";
                worksheet.Cell(row, 10).Value = "Flow_Debit";
                worksheet.Cell(row, 11).Value = "Remain_Last_Credit";
                worksheet.Cell(row, 12).Value = "Remain_last_Debit";
                //worksheet.Cell(row, 13).Value = "Account_Remain";

                row = 2;
                int writeValue = 0;

                foreach (var item in financialRecords)
                {
                    // بررسی گزینه همه رکوردها یا فقط مانده دار ها 
                    if (requestDB.AllOrHasMandeh == "2" && await Calculate_Last_Remain(item) && requestDB.GardeshOrMandeh == "1")
                    {
                        continue;
                    }
                    else
                    {
                        worksheet.Cell(row, 1).Value = item.Branch_ID;
                        worksheet.Cell(row, 2).Value = item.RBank_Code;
                        worksheet.Cell(row, 3).Value = item.RBank_Title;
                        worksheet.Cell(row, 4).Value = item.FinApplication_ID;
                        worksheet.Cell(row, 5).Value = item.FinApplication_Title;
                        worksheet.Cell(row, 6).Value = item.Motamam;
                        worksheet.Cell(row, 7).Value = item.Remain_First_Credit;
                        worksheet.Cell(row, 8).Value = item.Remain_First_Debit;
                        worksheet.Cell(row, 9).Value = item.Flow_Credit;
                        worksheet.Cell(row, 10).Value = item.Flow_Debit;
                        worksheet.Cell(row, 11).Value = item.Remain_Last_Credit;
                        worksheet.Cell(row, 12).Value = item.Remain_last_Debit;
                        //worksheet.Cell(row, 13).Value = item.Account_Remain;

                        row++;

                        writeValue++;
                    }
                }
                // اگر همه ستون دار ها بدون مانده باشد، هیچ رکوردی در اکسل نوشته نمیشود 
                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"BalanceGenerator:GenerateRawLGTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                // استایل دهی به اکسل گزارش دهی
                worksheet.Style.Font.FontName = "B Nazanin";
                worksheet.Style.Font.FontSize = 11;

                var range = worksheet.Range("G:M");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheet.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Column("C").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Column("E").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                var usedRange = worksheet.RangeUsed();

                if (usedRange != null)
                {
                    worksheet.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheet.Range("A1:L1");
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
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:GenerateRawLGTablesAsync --typeReport:Error");

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

        public async Task<bool> Calculate_Last_Remain(GLFinancialRecord Record)
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

