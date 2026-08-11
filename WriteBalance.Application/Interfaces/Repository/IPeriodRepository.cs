using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;

namespace WriteBalance.Application.Interfaces.Repository
{
    public interface IPeriodRepository
    {
        (int, bool, DateTime, DateTime) GetTimeAsync(APIRequestDto request, string FolderPath);
    }
}
