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
    public class BalanceCheck
    {
        public BalanceMerge _balanceMerge;
        public BalanceCheck(BalanceMerge balanceMerge)
        {
            _balanceMerge = balanceMerge;
        }


        //بررسی بالانس بودن و خالی نبودن شرح و یونیک بودن کد ها 
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
                    mergedRows = _balanceMerge.MergeDuplicateRows(mergedRows);
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
                        await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");
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
                        Logger.WriteEntry(JsonConvert.SerializeObject($"Not Balance with  {ekhtelaf}"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");
                        Logger.WriteEntry(JsonConvert.SerializeObject($"totalBed: {totalBed}, totalBes: {totalBes}"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");
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

                //  برای همه ترازها چک میکند که حداقل یک رکورد وجود داشته باشد که از کد 2 و کد 5 دارای مانده باشد 
                if (!mergedRows.Any(x => (x.Col1[0] == '2' && x.Col3 - x.Col4 != 0)) || !mergedRows.Any(x => (x.Col1[0] == '5' && x.Col3 - x.Col4 != 0)))
                {
                    throw new ConnectionMessageException(
                         new ConnectionMessage
                         {
                             MessageType = MessageType.Error,
                             Messages = new List<string> { $"کد 2 و 5 در این تراز وجود ندارد" }
                         },
                     requestDB.FolderPath
                     );
                }

                return await Task.FromResult(mergedRows);
            }
            catch (ConnectionMessageException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:checkBalance --typeReport:Error");
                throw;
            }
        }

        //بررسی بالانس بودن و خالی نبودن شرح و یونیک بودن کد ها 
        public async Task<List<ExcelRow>> checkGardeshBalance(List<ExcelRow> mergedRows, IExcelExporter excelExporter, DBRequestDto requestDB, MemoryStream streamReport)
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
                    mergedRows = _balanceMerge.MergeDuplicateGardeshRows(mergedRows);
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

                decimal totalBed = mergedRows.Sum(r => r.Col5??0);
                decimal totalBes = mergedRows.Sum(r => r.Col6??0);
                var ekhtelaf = totalBed - totalBes;

                if (totalBed != totalBes)
                {
                    if (Math.Abs(ekhtelaf) > 100)
                    {
                        await excelExporter.SaveReportAsync(streamReport, requestDB.FolderPath, $"گزارش {requestDB.FileName}");
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
                        Logger.WriteEntry(JsonConvert.SerializeObject($"Not Balance with  {ekhtelaf}"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");
                        Logger.WriteEntry(JsonConvert.SerializeObject($"totalBed: {totalBed}, totalBes: {totalBes}"), $"BalanceGenerator:GeneratePoyaTablesAsync --typeReport:Error");
                        if (totalBed > totalBes)
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "گردش بالانس",
                                Col5 = 0,
                                Col6 = Math.Abs(ekhtelaf),
                            });
                        }
                        else
                        {
                            mergedRows.Add(new ExcelRow
                            {
                                Col1 = "123456789",
                                Col2 = "گردش بالانس",
                                Col5 = Math.Abs(ekhtelaf),
                                Col6 = 0,
                            });
                        }
                    }
                }

                //  برای همه ترازهای گردش چک میکند که حداقل یک رکورد وجود داشته باشد که از کد 2 و کد 5 دارای مانده باشد 
                if (!mergedRows.Any(x => (x.Col1[0] == '2' && x.Col5 - x.Col6 != 0)) || !mergedRows.Any(x => (x.Col1[0] == '5' && x.Col5 - x.Col6 != 0)))
                {
                    throw new ConnectionMessageException(
                         new ConnectionMessage
                         {
                             MessageType = MessageType.Error,
                             Messages = new List<string> { $"کد 2 و 5 در این تراز وجود ندارد" }
                         },
                     requestDB.FolderPath
                     );
                }

                return await Task.FromResult(mergedRows);
            }
            catch (ConnectionMessageException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"BalanceGenerator:checkBalance --typeReport:Error");
                throw;
            }
        }
    
    }
}
