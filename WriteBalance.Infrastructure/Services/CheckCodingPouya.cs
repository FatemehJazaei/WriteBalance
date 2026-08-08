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
        public async Task<List<ExcelRow>> HandelNotFoundExcelAsync(List<ExcelRow> rows, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                var (notFound, excelRows) = ReplaceCol1Values(rows, requestDB);
                SaveNotFoundExcelAsync(notFound, excelExporter, requestDB);
                return rows;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CheckCodingPouya:SaveNotFoundExcelAsync --typeReport:Error");
                throw;
            }
        }

        public (List<string>, List<ExcelRow>) ReplaceCol1Values(List<ExcelRow> rows, DBRequestDto requestDB)
        {
            try
            {
                var notFound = new List<string>();

                foreach (var row in rows)
                {
                    if (requestDB.PouyaCodings.TryGetValue(row.Col1, out var value))
                    {
                        row.Col1 = value;
                    }
                    else
                    {
                        notFound.Add(row.Col1);
                    }
                }

                return (notFound, rows) ;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"CheckCodingPouya:ReplaceCol1Values --typeReport:Error");
                throw;
            }
        }

        public async Task SaveNotFoundExcelAsync(List<string> notFound, IExcelExporter excelExporter, DBRequestDto requestDB)
        {
            try
            {
                if (notFound == null || notFound.Count == 0)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject(" notFound.Count  = 0 or  notFound.Count  = null"), $"CheckCodingPouya:SaveNotFoundExcelAsync --typeReport:Error");
                    return;
                }
                   

                var workbookPouyaNotFound = excelExporter.GetWorkbookPouyaNotFound();
                var worksheet = workbookPouyaNotFound.Worksheets.Add("NotFound");

                // Header
                worksheet.Cell(1, 1).Value = "کد پیدا نشده";
                worksheet.Cell(1, 2).Value = "کد معادل  ";
                // Data
                for (int i = 0; i < notFound.Count; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = notFound[i];
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

                await excelExporter.SaveReportAsync(stream, requestDB.FolderPath, $" لیست کدهای معادل یافت نشده ");
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
