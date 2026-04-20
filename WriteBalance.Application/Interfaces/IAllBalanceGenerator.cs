using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public interface IAllBalanceGenerator
    {
        Task<List<ExcelRow>> CheckExcelRowGLAsync(List<FinancialRecord> financialRecords, DBRequestDto requestDB);
        Task<MemoryStream> GenerateAllTableAsync(List<FinancialRecord> financialRecords, List<ExcelRow> mergedRows, IExcelExporter excelExporter, DBRequestDto requestDB);
        Task<MemoryStream> GenerateAllTableGardeshAsync(List<FinancialRecord> financialRecords, List<ExcelRow> mergedRows, IExcelExporter excelExporter, DBRequestDto requestDB);
    }
}
