using Azure.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Infrastructure.Services
{
    public class CheckCodingPouya
    {
        public async Task<List<ExcelRow>> HandelNotFoundExcelAsync( List<ExcelRow> rows, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Start HandelNotFoundExcelAsync"), $"CheckCodingPouya:HandelNotFoundExcelAsync --typeReport:Debug");

                var (notFound, excelRows) = ReplaceCol1Values(rows, requestDB);
                await SaveNotFoundExcelAsync(notFound, excelExporter, requestDB);
                return excelRows;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CheckCodingPouya:SaveNotFoundExcelAsync --typeReport:Error");
                throw;
            }
        }

        public (List<EquivalentCodePouya>, List<ExcelRow>) ReplaceCol1Values(
            List<ExcelRow> rows,
            DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(
                    JsonConvert.SerializeObject("Start ReplaceCol1Values"),
                    "CheckCodingPouya:ReplaceCol1Values --typeReport:Debug");

                var notFound = new List<EquivalentCodePouya>();

                if (requestDB.PouyaCodings == null ||
                    requestDB.PouyaCodings.Count == 0)
                {
                    notFound.AddRange(
                        rows.Select(x => new EquivalentCodePouya
                        {
                            SourceCode = x.Col1,
                            EquivalentCode = string.Empty
                        }));

                    return (notFound, rows);
                }

                // فقط یک بار ساخته می‌شود
                var codingDictionary = requestDB.PouyaCodings
                    .ToDictionary(x => x.SourceCode);

                // هر ExcelRow فقط یک lookup انجام می‌دهد
                foreach (var row in rows)
                {
                    var sourceCode = row.Col1;

                    if (codingDictionary.TryGetValue(sourceCode, out var coding))
                    {
                        row.Col1 = coding.EquivalentCode;
                    }
                    else
                    {
                        notFound.Add(new EquivalentCodePouya
                        {
                            SourceCode = sourceCode,
                            EquivalentCode = string.Empty
                        });
                    }
                }

                return (notFound, rows);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(
                    JsonConvert.SerializeObject(ex),
                    "CheckCodingPouya:ReplaceCol1Values --typeReport:Error");

                throw;
            }
        }


        //public (List<string>, List<ExcelRow>) ReplaceCol1Values(List<ExcelRow> rows, DBRequestDto requestDB)
        //{
        //    try
        //    {
        //        Logger.WriteEntry(JsonConvert.SerializeObject("Start ReplaceCol1Values"), $"CheckCodingPouya:ReplaceCol1Values --typeReport:Debug");
        //        var notFound = new List<string>();

        //        foreach (var row in rows)
        //        {
        //            if (requestDB.PouyaCodings != null &&  (requestDB.PouyaCodings.TryGetValue(row.Col1, out var value)))
        //            {
        //                row.Col1 = value;
        //            }
        //            else
        //            {
        //                notFound.Add(row.Col1);
        //            }
        //        }

        //        return (notFound, rows) ;
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CheckCodingPouya:ReplaceCol1Values --typeReport:Error");
        //        throw;
        //    }
        //}

        public async Task SaveNotFoundExcelAsync( List<EquivalentCodePouya> notFound, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Start SaveNotFoundExcelAsync"), $"CheckCodingPouya:SaveNotFoundExcelAsync --typeReport:Debug");

                if (notFound == null || notFound.Count == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject(" notFound.Count  = 0 or  notFound.Count  = null"), $"CheckCodingPouya:SaveNotFoundExcelAsync --typeReport:Error");

                    return;
                }

                Logger.WriteEntry(JsonConvert.SerializeObject($"notFound.Count {notFound.Count}"), $"CheckCodingPouya:SaveNotFoundExcelAsync --typeReport:Error");
                var workbookPouyaNotFound = excelExporter.GetWorkbookPouyaNotFound();
                var worksheet = workbookPouyaNotFound.Worksheets.Add("NotFound");
                worksheet.RightToLeft = true;

                // Header
                worksheet.Cell(1, 1).Value = "کد پیدا نشده";
                worksheet.Cell(1, 2).Value = "کد معادل";

                // Data
                for (int i = 0; i < notFound.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value =
                        notFound[i]?.SourceCode ?? "";

                    worksheet.Cell(i + 2, 2).Value =
                        notFound[i]?.EquivalentCode ?? "";
                }

                //استایل دهی  گزارش
                worksheet.Style.Font.FontName = "B Nazanin";
                worksheet.Style.Font.FontSize = 11;

                var range = worksheet.Range("A:B");
                range.Style.NumberFormat.Format = "#,##0_);[Red](#,##0)";

                worksheet.Columns().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var usedRange = worksheet.RangeUsed();

                if (usedRange != null)
                {
                    worksheet.Columns().AdjustToContents();
                    usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    usedRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                var headerRange = worksheet.Range("A1:B1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LapisLazuli;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Font.FontColor = XLColor.White;

                //ذخیره فایل گزارش
                var stream = new MemoryStream();
                workbookPouyaNotFound.SaveAs(stream);
                stream.Position = 0;

                await excelExporter.SavePouyaNotFoundAsync(stream, requestDB.FolderPath, $" لیست کدهای معادل یافت نشده {requestDB.FileName}");
                return;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CheckCodingPouya:SaveNotFoundExcelAsync --typeReport:Error");
                throw;
            }
           
        }

    }
}
