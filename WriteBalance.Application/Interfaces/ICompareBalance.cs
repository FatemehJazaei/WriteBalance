using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public interface ICompareBalance
    {
          Task<List<ExcelRow>> SetExcelRowAsync(List<FinancialRecord> financialRecords);
          Task<List<ExcelRow>> SetGLExcelRowAsync(List<GLFinancialRecord> financialRecords);
          Task<List<ExcelRow>> SetRayanExcelRowAsync(List<RayanFinancialRecord> financialRecords);
          Task<List<ExcelRow>> SetPouyaExcelRowAsync(List<PouyaFinancialRecord> financialRecords);
          Task<List<ExcelRow>> CreateAllExcelRowAsync(List<ExcelRow> samaExcelRow, List<ExcelRow> hamrahExcelRow, List<ExcelRow> karbourdiExcelRow, List<ExcelRow> rayanExcelRow, List<ExcelRow> pouyaExcelRow);
          Task<List<CompareRows>> CompareBalanceAsync(List<ExcelRow> allFinancialRecords, List<ExcelRow> GLFinancialRecords,  DBRequestDto requestDB);
          Task WriteExcelAsync(List<CompareRows> ExcelRows, IExcelExporter excelExporter, DBRequestDto requestDB);
    }
}
