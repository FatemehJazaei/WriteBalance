using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public interface IPouyaBalanceGenerator
    {
        Task<(MemoryStream, MemoryStream)> GeneratePoyaTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
        Task<(MemoryStream, MemoryStream)> GeneratePoyaGardeshTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
        
    }
}
