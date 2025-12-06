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
        List<FinancialRecord> ExecuteSPList(APIRequestDto request,  DBRequestDto requestDB, string startTime , string endTime);
        List<RayanFinancialRecord> ExecuteRayanSPList(APIRequestDto request, DBRequestDto requestDB, string startTime, string endTime);
        List<PouyaFinancialRecord> ExecutePoyaSPList(APIRequestDto request, DBRequestDto requestDB, string startTime, string endTime);
    }
}
