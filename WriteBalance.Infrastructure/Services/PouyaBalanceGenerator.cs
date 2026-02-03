using ClosedXML.Excel;
using DocumentFormat.OpenXml;
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
    public class PouyaBalanceGenerator: IPouyaBalanceGenerator
    {
        public BalanceMerge _balanceMerge;
        public BalanceCheck _balanceCheck;
        public PouyaBalanceGenerator(BalanceMerge balanceMerge, BalanceCheck balanceCheck)
        {
            _balanceMerge = balanceMerge;
            _balanceCheck = balanceCheck;
        }


        // تولید جدول برای پویا
        public async Task<(MemoryStream, MemoryStream)> GeneratePoyaTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GeneratePoyaTablesAsync"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Info");
                //برای تراز پویا  دو فایل تراز ریالی و ارزی  و یک فایل گزارش داریم 
                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();
                var workbookUploadArzi = excelExporter.GetWorkbookUploadArzi();

                //  تولید جدول تراز خام برای گزارش دهی
                var streamReport = await GenerateRawPouyaTablesAsync(financialRecords, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                // فرآیند تولید تراز ریالی 
                //حذف کد 6
                var rowsRial = financialRecords
                .Where(x => (x.Kol_Code.ToString() != null && x.Kol_Code.ToString()[0] != '6'))
                .Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}_{x.Arz_Code}_{x.Moeen_Code}_{x.Code_Arz_Abbr}",
                    Col2 = $"{x.Kol_Title}_{x.Sharh_Arz}",
                    Col3 = x.Mande_Bed_rial ?? 0,
                    Col4 = x.Mande_Bes_rial ?? 0,
                }).ToList();

                //یونیک کردن کدها
                var mergedRows = _balanceMerge.MergeDuplicateRows(rowsRial);
                // بررسی بالانس بودن  تراز 
                mergedRows = await _balanceCheck.checkBalance(mergedRows, excelExporter, requestDB, streamReport);

                // اضافه کردن شیت برای تراز محاسبه شده اکسیر برای اپلود و گزارش دهی 
                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر ریالی");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
                int row = 2;
                int writeValue = 0;

                foreach (var item in mergedRows)
                {
                    //بررسی گزینه همه یا مانده دارها
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

                //بررسی اینکه همه رکورد ها مانده دار است یا نه 
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

                // ذخیره فایل اکسل گزارش
                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;
                

                var rowsArzi = financialRecords
                    .Where(x => (x.Kol_Code.ToString() != null && x.Kol_Code.ToString()[0] != '6'))
                    .Select(x => new ExcelRow
                    {
                        Col1 = $"{x.Kol_Code}_{x.Arz_Code}_{x.Moeen_Code}_{x.Code_Arz_Abbr}",
                        Col2 = $"{x.Kol_Title}_{x.Sharh_Arz}",
                        Col3 = x.Mande_Bed_arzi ?? 0,
                        Col4 = x.Mande_Bes_arzi ?? 0,
                    }).ToList();

                //یونیک کردن کدها
                var mergedRowsArzi = _balanceMerge.MergeDuplicateRows(rowsArzi);
                // بررسی بالانس بودن  تراز 
                mergedRowsArzi = await _balanceCheck.checkBalance(mergedRowsArzi, excelExporter, requestDB, streamReport);


                // شروع فرایند تولید اکسل ارزی
                var worksheetUploadArzi = workbookUploadArzi.Worksheets.Add("Data");
                var worksheetReportArzi = workbookReport.Worksheets.Add("تراز اکسیر ارزی");
                worksheetUploadArzi.RightToLeft = true;
                worksheetReportArzi.RightToLeft = true;
                row = 2;
                writeValue = 0;

                //فرایند نوشتن رکوردها
                foreach (var item in mergedRowsArzi)
                {
                    // چک کردن گزینه همه یا فقط مانده دارها 
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
                // بررسی حالتی که همه رکورد ها بدون مانده است
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

                // استایل دهی به گزارش ها
                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;

                range = worksheetReport.Range("C:D");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                // ذخیره فایل گزارش
                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUploadArzi = new MemoryStream();
                workbookUploadArzi.SaveAs(streamUploadArzi);
                streamUploadArzi.Position = 0;
                return (streamUpload, streamUploadArzi);

            }
            catch (ConnectionMessageException)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("GeneratePoyaTablesAsync failed!"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");
                throw;
            }
        }


        public async Task<(MemoryStream, MemoryStream)> GeneratePoyaGardeshTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GeneratePoyaTablesAsync"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Info");
                //برای تراز پویا  دو فایل تراز ریالی و ارزی  و یک فایل گزارش داریم 
                var workbookReport = excelExporter.GetWorkbookReport();
                var workbookUpload = excelExporter.GetWorkbookUpload();
                var workbookUploadArzi = excelExporter.GetWorkbookUploadArzi();

                //  تولید جدول تراز خام برای گزارش دهی
                var streamReport = await GenerateRawPouyaTablesAsync(financialRecords, excelExporter, workbookReport, requestDB);
                streamReport.Position = 0;

                // فرآیند تولید تراز ریالی 
                //حذف کد 6
                var rowsRial = financialRecords
                .Where(x => (x.Kol_Code.ToString() != null && x.Kol_Code.ToString()[0] != '6'))
                .Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}_{x.Arz_Code}_{x.Moeen_Code}_{x.Code_Arz_Abbr}",
                    Col2 = $"{x.Kol_Title}_{x.Sharh_Arz}",
                    Col3 = x.Mande_Bed_rial ?? 0,
                    Col4 = x.Mande_Bes_rial ?? 0,
                    Col5 = x.Gardersh_Bed_rial ?? 0,
                    Col6 = x.Gardersh_Bes_rial ?? 0,
                }).ToList();

                //یونیک کردن کدها
                var mergedRows = _balanceMerge.MergeDuplicateGardeshPouyaRows(rowsRial);
                // بررسی بالانس بودن  تراز 
                mergedRows = await _balanceCheck.checkGardeshBalance(mergedRows, excelExporter, requestDB, streamReport);

                // اضافه کردن شیت برای تراز محاسبه شده اکسیر برای اپلود و گزارش دهی 
                var worksheetUpload = workbookUpload.Worksheets.Add("Data");
                var worksheetReport = workbookReport.Worksheets.Add("تراز اکسیر ریالی");
                worksheetUpload.RightToLeft = true;
                worksheetReport.RightToLeft = true;
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
                    worksheetReport.Cell(row, 4).Value = item.Col6; ;

                    row++;
                    writeValue++;
                }

                //بررسی اینکه همه رکورد ها مانده دار است یا نه 
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

                // ذخیره فایل اکسل گزارش
                var streamUpload = new MemoryStream();
                workbookUpload.SaveAs(streamUpload);
                streamUpload.Position = 0;

                var rowsArzi = financialRecords
                    .Where(x => (x.Kol_Code.ToString() != null && x.Kol_Code.ToString()[0] != '6'))
                    .Select(x => new ExcelRow
                    {
                        Col1 = $"{x.Kol_Code}_{x.Arz_Code}_{x.Moeen_Code}_{x.Code_Arz_Abbr}",
                        Col2 = $"{x.Kol_Title}_{x.Sharh_Arz}",
                        Col3 = x.Mande_Bed_arzi ?? 0,
                        Col4 = x.Mande_Bes_arzi ?? 0,
                        Col5 = x.Gardersh_Bed_arzi ?? 0,
                        Col6 = x.Gardersh_Bes_arzi ?? 0,
                    }).ToList();


                //یونیک کردن کدها
                var mergedRowsArzi = _balanceMerge.MergeDuplicateGardeshPouyaRows(rowsArzi);
                // بررسی بالانس بودن  تراز 
                mergedRowsArzi = await _balanceCheck.checkGardeshBalance(mergedRowsArzi, excelExporter, requestDB, streamReport);

                // شروع فرایند تولید اکسل ارزی
                var worksheetUploadArzi = workbookUploadArzi.Worksheets.Add("Data");
                var worksheetReportArzi = workbookReport.Worksheets.Add("تراز اکسیر ارزی");
                worksheetUploadArzi.RightToLeft = true;
                worksheetReportArzi.RightToLeft = true;
                row = 2;
                writeValue = 0;

                //فرایند نوشتن رکوردها
                foreach (var item in mergedRowsArzi)
                {
                    worksheetUploadArzi.Cell(row, 1).Value = item.Col1;
                    worksheetUploadArzi.Cell(row, 2).Value = item.Col2;
                    worksheetUploadArzi.Cell(row, 3).Value = item.Col5.ToString();
                    worksheetUploadArzi.Cell(row, 4).Value = item.Col6.ToString();

                    worksheetReportArzi.Cell(row, 1).Value = item.Col1;
                    worksheetReportArzi.Cell(row, 2).Value = item.Col2;
                    worksheetReportArzi.Cell(row, 3).Value = item.Col5;
                    worksheetReportArzi.Cell(row, 4).Value = item.Col6;

                    row++;
                    writeValue++;
                }
                // بررسی حالتی که همه رکورد ها بدون مانده است
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

                range = worksheetReport.Range("C:D");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheetReport.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheetReport.Column("B").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                usedRange = worksheetReport.RangeUsed();

                if (usedRange != null)
                {
                    worksheetReport.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }
                streamReport.Position = 0;

                // ذخیره فایل گزارش
                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");

                var streamUploadArzi = new MemoryStream();
                workbookUploadArzi.SaveAs(streamUploadArzi);
                streamUploadArzi.Position = 0;
                return (streamUpload, streamUploadArzi);

            }
            catch (ConnectionMessageException ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"GeneratePoyaTablesAsync failed! {ex}"), $"BalanceGenerator:GenerateTablesAsync --typeReport:Error");
                throw;
            }
        }

        //تولید تراز خام پویا
        public async Task<MemoryStream> GenerateRawPouyaTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, XLWorkbook workbook, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GenerateRawPouyaTablesAsync"), $"BalanceGenerator:GenerateRawPouyaTablesAsync --typeReport:Info");
                var worksheet = workbook.Worksheets.Add("تراز خام");
                worksheet.RightToLeft = true;

                int row = 1;

                // مقدار دهی سطر اول برای گزارش 
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

                // شروع نوشتن تراز ها
                foreach (var item in financialRecords)
                {
                    // بررسی گزینه همه یا فقط مانده دار
                    if (requestDB.AllOrHasMandeh == "2" && item.Mande_Bed_arzi - item.Mande_Bes_arzi == 0 && requestDB.GardeshOrMandeh == "1")
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

                //استایل دهی  گزارش
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

                //ذخیره فایل گزارش
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

    }
}
