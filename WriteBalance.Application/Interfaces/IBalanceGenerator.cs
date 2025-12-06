using DocumentFormat.OpenXml.Office2016.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public  interface IBalanceGenerator
    {
        Task<MemoryStream> GenerateTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
        Task<MemoryStream> GenerateRayanTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
        Task GeneratePoyaTablesAsync(List<PouyaFinancialRecord> financialRecords, IExcelExporter excelExporter, DBRequestDto requestDB);
    }
}
