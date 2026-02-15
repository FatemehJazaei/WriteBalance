using ClosedXML.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;
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
using WriteBalance.Domain.Entities;

namespace WriteBalance.Infrastructure.Services
{
    public class RayanBalanceGenerator : IRayanBalanceGenerator
    {
        public BalanceMerge _balanceMerge;
        public BalanceCheck _balanceCheck;
        public RayanBalanceGenerator(BalanceMerge balanceMerge, BalanceCheck balanceCheck)
        {
            _balanceMerge = balanceMerge;
            _balanceCheck = balanceCheck;
        }


        //تولید تراز  رایان
        public async Task<MemoryStream> GenerateRayanTablesAsync(List<RayanFinancialRecord> RayanFinancialRecord, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRayanTablesAsync"), $"RayanBalanceGenerator:GenerateRayanTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();

                //تولید تراز خام برای گزارش
                var streamReport = await GenerateRawRayanTablesAsync(RayanFinancialRecord, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                //فیلتر کد کل 6
                var filteredSource = RayanFinancialRecord
                    .Where(x => x.Kol_Code[0] != '6');

                //حذف کدهای دریافتی از کاربر تا سطح معین 
                List<ExcelRow> specialRows = new List<ExcelRow>();
                if (requestDB.ExceptCode.Count != 0)
                {
                    specialRows = filteredSource
                    .Where(x => requestDB.ExceptCode.Any(ec =>
                        ec.Kol_Code == x.Kol_Code &&
                        ec.Moeen_Code == x.Moeen_Code[^3..]))
                    .GroupBy(x => new { x.Kol_Code, x.Moeen_Code })
                    .Select(g =>
                    {
                        var first = g.First();

                        return new ExcelRow
                        {
                            Col1 = $"{first.Kol_Code}_{first.Moeen_Code[^3..]}_0_0_0",
                            Col2 = $"{first.Kol_Title}_{first.Moeen_Title}",
                            Col3 = g.Sum(x => (decimal)x.Mande_Bed),
                            Col4 = g.Sum(x => (decimal)x.Mande_Bes)
                        };
                    })
                    .ToList();
                }

                //  بقیه رکوردها   
                var normalRows = filteredSource
                    .Where(x => !requestDB.ExceptCode.Any(ec =>
                        ec.Kol_Code == x.Kol_Code &&
                        ec.Moeen_Code == x.Moeen_Code[^3..]))
                    .Select(x =>
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
                            code += "_0";
                        }

                        if (x.joze2_Code.Length == 21)
                        {
                            code += $"_{x.joze2_Code[^4..]}";
                            title += $"_{x.joze2_Title}";
                        }
                        else
                        {
                            code += "_0";
                        }

                        return new ExcelRow
                        {
                            Col1 = code,
                            Col2 = title,
                            Col3 = (decimal)x.Mande_Bed,
                            Col4 = (decimal)x.Mande_Bes
                        };
                    });


                //  ترکیب نهایی
                var rows = specialRows
                    .Concat(normalRows)
                    .ToList();

                // مرج کدها 
                var mergedRows = _balanceMerge.MergeDuplicateRows(rows);
                // بررسی بالانس بودن  تراز 
                mergedRows = await _balanceCheck.checkBalance(mergedRows, excelExporter, requestDB, streamReport);

                var worksheetUpload = workbookUpload.Worksheets.Add("data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                //سطر اول خالی است
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {
                    // AllOrHasMandeh : مقدار 1 همه رکورد ها را برمیگرداند و مقدار 2 فقط مانده دار ها 
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

                // بررسی اینکه همه رکوردها بدون مانده است
                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"RayanBalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                // استایل دهی به گزارش ها
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

                //ذخیره فایل گزارش  
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
                Logger.WriteEntry(JsonConvert.SerializeObject("Failed to GenerateRayanTablesAsync"), $"RayanBalanceGenerator:GenerateRayanTablesAsync --typeReport:Error");
                throw;
            }
        }


        //تولید تراز گردش رایان
        public async Task<MemoryStream> GenerateRayanGardeshTablesAsync(List<RayanFinancialRecord> RayanFinancialRecord, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRayanTablesAsync"), $"RayanBalanceGenerator:GenerateRayanTablesAsync --typeReport:Info");

                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();

                //تولید تراز خام برای گزارش
                var streamReport = await GenerateRawRayanTablesAsync(RayanFinancialRecord, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                //فیلتر کد کل 6 
                var rows = RayanFinancialRecord
                    .Where(x => (x.Kol_Code != null && x.Kol_Code[0] != '6'))
                    .Select(x =>
                    {
                        var code = $"{x.Kol_Code}_{x.Moeen_Code[^3..]}";
                        var title = $"{x.Kol_Title}_{x.Moeen_Title}";

                        return new ExcelRow
                        {
                            Col1 = code,
                            Col2 = title,
                            Col3 = (decimal)x.Mande_Bed,
                            Col4 = (decimal)x.Mande_Bes,
                            Col5 = (decimal)x.bedehkar,
                            Col6 = (decimal)x.bestankar,
                        };
                    }).ToList();

                // مرج کدها 
                var mergedRows = _balanceMerge.MergeDuplicateGardeshRows(rows);

                // بررسی بالانس بودن  تراز 
                mergedRows = await _balanceCheck.checkGardeshBalance(mergedRows, excelExporter, requestDB, streamReport);

                var worksheetUpload = workbookUpload.Worksheets.Add("data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                //سطر اول خالی است
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

                // بررسی اینکه همه رکوردها بدون مانده است
                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"RayanBalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"تمام سطرها بدون مانده میباشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                // استایل دهی به گزارش ها
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

                //ذخیره فایل گزارش  
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
                Logger.WriteEntry(JsonConvert.SerializeObject($"Failed to GenerateRayanTablesAsync {ex}"), $"RayanBalanceGenerator:GenerateRayanTablesAsync --typeReport:Error");
                throw;
            }
        }

        // تولید جدول خام تراز رایان
        public async Task<MemoryStream> GenerateRawRayanTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawRayanTablesAsync"), $"RayanBalanceGenerator:GenerateRawRayanTablesAsync --typeReport:Info");
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
                    //بررسی گزینه همه یا فقط مانده دارها
                    if (requestDB.AllOrHasMandeh == "2" && item.Mande_Bed - item.Mande_Bes == 0 && requestDB.GardeshOrMandeh == "1")
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

                //چک کردن همه رکوردها بدون مانده باشد
                if (writeValue == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"All records dont have mande."), $"RayanBalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");

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

                // ذخیره
                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                stream.Position = 0;
                return await Task.FromResult(stream);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"RayanBalanceGenerator:GenerateRawRayanTablesAsync --typeReport:Error");

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