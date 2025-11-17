using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;

namespace WriteBalance.Application.Interfaces
{
    public interface IPeriodRepository
    {
        Task<(int, DateTime, DateTime)> GetTimeAsync(APIRequestDto request);
    }
}
