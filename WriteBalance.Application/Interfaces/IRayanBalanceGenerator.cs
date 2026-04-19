using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public interface IRayanBalanceGenerator
    {
        Task<MemoryStream> GenerateRayanTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
        Task<MemoryStream> GenerateRayanGardeshTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
        List<RayanFinancialRecord> ExceptRayanTables(List<RayanFinancialRecord> RayanFinancialRecord, DBRequestDto requestDB);
    }
}
