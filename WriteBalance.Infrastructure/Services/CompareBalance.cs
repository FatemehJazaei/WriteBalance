using Azure.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
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
    public class CompareBalance : ICompareBalance
    {
        public BalanceMerge _balanceMerge;
        public BalanceCheck _balanceCheck;
        public CalculateNewRows _calculateNewRows;
        public CompareBalance(BalanceMerge balanceMerge, BalanceCheck balanceCheck, CalculateNewRows calculateNewRows)
        {
            _balanceMerge = balanceMerge;
            _balanceCheck = balanceCheck;
            _calculateNewRows = calculateNewRows;
        }

  
        public async Task<List<ExcelRow>> SetExcelRowAsync(List<FinancialRecord> financialRecords)
        {
            try
            {
                var rows = financialRecords
                .Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}",
                    Col2 = $"{x.Kol_Title}",
                    Col3 = x.Remain_First_Debit ?? decimal.Zero,
                    Col4 = x.Remain_First_Credit ?? decimal.Zero,
                    Col5 = x.Flow_Debit ?? decimal.Zero,
                    Col6 = x.Flow_Credit ?? decimal.Zero,
                }).ToList();

                // محاسبه مانده برای هر رکورد
                var rowsEditRemain = await _calculateNewRows.Calculate_New_rows(rows);
                // یونیک کردن رکوردها بر اساس کد
                var mergedRows = _balanceMerge.MergeDuplicateRows(rowsEditRemain);

                return await Task.FromResult(mergedRows);
            }
            catch (ConnectionMessageException ex)
            {
                throw;
            }
            catch (Exception ex) 
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CompareBalance:SetExcelRowAsync --typeReport:Error");
                throw;
            }
        }

        public async Task<List<ExcelRow>> SetGLExcelRowAsync(List<GLFinancialRecord> financialRecords)
        {
            try
            {
                var rows = financialRecords
                    .Select(x => new ExcelRow
                    {
                        Col1 = $"{x.RBank_Code[..4]}",
                        Col2 = $"{x.RBank_Title.Replace("***", "_")}",
                        Col3 = x.Remain_last_Debit ?? decimal.Zero,
                        Col4 = x.Remain_Last_Credit ?? decimal.Zero,
                        Col5 = x.Flow_Debit ?? decimal.Zero,
                        Col6 = x.Flow_Credit ?? decimal.Zero,
                    }).ToList();

                // محاسبه مانده برای هر رکورد
                var rowsEditRemain = await _calculateNewRows.Calculate_New_rows(rows);
                // یونیک کردن رکوردها بر اساس کد
                var mergedRows = _balanceMerge.MergeDuplicateRows(rowsEditRemain);
                return await Task.FromResult(mergedRows);
            }
            catch (ConnectionMessageException ex)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CompareBalance:SetGLExcelRowAsync --typeReport:Error");
                throw;
            }
        }

        public async Task<List<ExcelRow>> SetRayanExcelRowAsync(List<RayanFinancialRecord> financialRecords)
        {
            try
            {
                var rows = financialRecords
                .Select(x =>
                     new ExcelRow
                     {
                         Col1 = x.Kol_Code,
                         Col2 = x.Kol_Title,
                         Col3 = (decimal)x.Mande_Bed,
                         Col4 = (decimal)x.Mande_Bes
                     }
                ).ToList();

                var mergedRows = _balanceMerge.MergeDuplicateRows(rows);

                return await Task.FromResult(mergedRows);
            }
            catch (ConnectionMessageException ex)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CompareBalance:SetRayanExcelRowAsync --typeReport:Error");
                throw;
            }
        }

        public async Task<List<ExcelRow>> SetPouyaExcelRowAsync(List<PouyaFinancialRecord> financialRecords)
        {
            try
            {
                var rowsRial = financialRecords
                .Select(x => new ExcelRow
                {
                    Col1 = $"{x.Kol_Code}",
                    Col2 = $"{x.Kol_Title}",
                    Col3 = x.Mande_Bed_rial ?? 0,
                    Col4 = x.Mande_Bes_rial ?? 0,
                }).ToList();

                //یونیک کردن کدها
                var mergedRows = _balanceMerge.MergeDuplicateRows(rowsRial);

                return await Task.FromResult(mergedRows);
            }
            catch (ConnectionMessageException ex)
            {             
                throw;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CompareBalance:SetPouyaExcelRowAsync --typeReport:Error");
                throw;
            }

        }

        public async Task<List<ExcelRow>> CreateAllExcelRowAsync(List<ExcelRow> samaExcelRow, List<ExcelRow> hamrahExcelRow, List<ExcelRow> karbourdiExcelRow, List<ExcelRow> rayanExcelRow, List<ExcelRow> pouyaExcelRow)
        {
            try
            {
                var allExcelRows = samaExcelRow.Concat(hamrahExcelRow).ToList();
                allExcelRows = allExcelRows.Concat(karbourdiExcelRow).ToList();
                allExcelRows = allExcelRows.Concat(rayanExcelRow).ToList();
                allExcelRows = allExcelRows.Concat(pouyaExcelRow).ToList();

                var mergedRows = _balanceMerge.MergeDuplicateRows(allExcelRows);
                return await Task.FromResult(mergedRows);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CompareBalance:CreateAllExcelRowAsync --typeReport:Error");
                throw;
            }
        }


        public async Task<List<CompareRows>> CompareBalanceAsync(List<ExcelRow> allExcelRows, List<ExcelRow> GLExcelRows,  DBRequestDto requestDB)
        {
            try
            {
                var dictGL = GLExcelRows.ToDictionary(x => x.Col1);
                var dictAll = allExcelRows.ToDictionary(x => x.Col1);

                var result = new List<CompareRows>();

                // بررسی آیتم‌های GLExcelRows
                foreach (var a in GLExcelRows)
                {
                    if (dictAll.TryGetValue(a.Col1, out var b))
                    {
                        // اگر در هر دو وجود دارد ولی ستون 3 یا 4 متفاوت است
                        if (a.Col3 != b.Col3 || a.Col4 != b.Col4)
                        {
                            result.Add(new CompareRows
                            {
                                Code = a.Col1,
                                Titel = a.Col2,
                                MandehGL = (decimal)(a.Col3 - a.Col4),
                                MnadehAll = (decimal)(b.Col3 - b.Col4),
                                Ekhtelaf = (decimal)Math.Abs((a.Col3 - a.Col4) - (b.Col3 - b.Col4))
                            });
                        }
                    }
                    else
                    {
                        // اگر فقط در GLExcelRows وجود دارد
                        result.Add(new CompareRows
                        {
                            Code = a.Col1,
                            Titel = a.Col2,
                            MandehGL = (decimal)(a.Col3 - a.Col4),
                            MnadehAll = null,
                            Ekhtelaf = null
                        });
                    }
                }

                // بررسی آیتم‌هایی که فقط در listB هستند
                foreach (var b in allExcelRows)
                {
                    if (!dictGL.ContainsKey(b.Col1))
                    {
                        result.Add(new CompareRows
                        {
                            Code = b.Col1,
                            Titel = b.Col2,
                            MandehGL = null,
                            MnadehAll = (decimal)(b.Col3 - b.Col4),
                            Ekhtelaf = null
                        });
                    }
                }

                if (allExcelRows.Count == 0)
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { "مقایسه با موفقیت انجام شد. مغایرتی یافت نشد." }
                        },
                    requestDB.FolderPath
                    );
                }

                return await Task.FromResult(result);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CompareBalance:CompareBalanceAsync --typeReport:Error");
                throw;
            }
        }

        public async Task WriteExcelAsync(List<CompareRows> ExcelRows, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                var workbookReport = excelExporter.GetWorkbookReport();
                var worksheetReport = workbookReport.Worksheets.Add("نتیجه مقایسه");
                worksheetReport.RightToLeft = true;

                int row = 1;
                worksheetReport.Cell(row, 1).Value = "کد کل";
                worksheetReport.Cell(row, 2).Value = "شرح";
                worksheetReport.Cell(row, 3).Value = "مانده ترازها";
                worksheetReport.Cell(row, 4).Value = "مانده تراز جی ال";
                worksheetReport.Cell(row, 5).Value = " اختلاف";
                row = 2;

                foreach (var item in ExcelRows)
                {
                    worksheetReport.Cell(row, 1).Value = item.Code;
                    worksheetReport.Cell(row, 2).Value = item.Titel;
                    worksheetReport.Cell(row, 3).Value = item.MnadehAll;
                    worksheetReport.Cell(row, 4).Value = item.MandehGL;
                    worksheetReport.Cell(row, 5).Value = item.Ekhtelaf;

                    row++;
                }

                //استایل دهی به گزارش 
                worksheetReport.Style.Font.FontName = "B Nazanin";
                worksheetReport.Style.Font.FontSize = 11;

                var range = worksheetReport.Range("C:E");
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

                var headerRange = worksheetReport.Range("A1:E1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;

                var streamReport = new MemoryStream();
                workbookReport.SaveAs(streamReport);
                streamReport.Position = 0;
                await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, requestDB.FileName);

                return;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CompareBalance:WriteExcelAsync --typeReport:Error");
                throw;
            }
           
        }

    }
}
