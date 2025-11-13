using DocumentFormat.OpenXml.Office2016.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public  interface IBalanceGenerator
    {
        Task<MemoryStream> GenerateTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, string FolderPath);
        Task<MemoryStream> GenerateRayanTablesAsync(List<RayanFinancialRecord> financialRecords, IExcelExporter excelExporter, string FolderPath);
        Task<MemoryStream> GeneratePoyaTablesAsync(List<FinancialRecord> financialRecords, IExcelExporter excelExporter, string FolderPath);
    }
}
