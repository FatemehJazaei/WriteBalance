using Azure.Core;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;
using WriteBalance.Infrastructure.Context;

namespace WriteBalance.Infrastructure.Repositories
{
    public class PeriodRepository : IPeriodRepository
    {
        private readonly AppDbContext _context;

        public PeriodRepository(AppDbContext context)
        {
            _context = context;
        }

        public async Task<(int, DateTime, DateTime)> GetTimeAsync(APIRequestDto request)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GetTimeAsync"), $"PeriodRepository:GetTimeAsync--typeReport:Info");

                var entity = await _context.Periods
                    .AsNoTracking()
                    .FirstOrDefaultAsync(x => x.Id == request.PeriodId);

                Logger.WriteEntry(JsonConvert.SerializeObject($"CompanyId:{entity.CompanyId},StartDate:{entity.StartDate},TimeEnd:{entity.TimeEnd} "), $"PeriodRepository:GetTimeAsync--typeReport:Debug");

                return (entity.CompanyId, entity.StartDate, entity.TimeEnd);
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"PeriodRepository:GetTimeAsync--typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject($"Failed to fetch StartDate and TimeEnd from database (periodId): {request.PeriodId}"), $"PeriodRepository:GetTimeAsync--typeReport:Error");

                throw new ConnectionMessageException(new ConnectionMessage
                {
                    MessageType = MessageType.Error,
                    Messages = new List<string> { "ارتباط با پایگاه داده اکسیر ناموفق!" }
                },
                request.FolderPath
                );
            }
        }
    }
}
