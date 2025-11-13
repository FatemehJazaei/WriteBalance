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
        List<FinancialRecord> ExecuteSPList(DBRequestDto requestDB);
        List<RayanFinancialRecord> ExecuteRayanSPList(DBRequestDto requestDB);
        List<FinancialRecord> ExecutePoyaSPList(DBRequestDto requestDB);
    }
}
