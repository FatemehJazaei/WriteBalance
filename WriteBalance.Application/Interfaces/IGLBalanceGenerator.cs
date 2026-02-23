using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public interface IGLBalanceGenerator
    {
        Task<MemoryStream> GenerateGLTablesAsync(List<GLFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
        Task<MemoryStream> GenerateGardeshGLTablesAsync(List<GLFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);


    }
}
