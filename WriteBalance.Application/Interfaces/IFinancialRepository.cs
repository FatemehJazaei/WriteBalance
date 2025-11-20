using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public  interface IFinancialRepository
    {
        List<FinancialRecord> ExecuteSPList(APIRequestDto request,  DBRequestDto requestDB, DateTime startTime , DateTime endTime);
        List<RayanFinancialRecord> ExecuteRayanSPList(APIRequestDto request, DBRequestDto requestDB, DateTime startTime, DateTime endTime);
        List<FinancialRecord> ExecutePoyaSPList(APIRequestDto request, DBRequestDto requestDB, DateTime startTime, DateTime endTime);
    }
}
