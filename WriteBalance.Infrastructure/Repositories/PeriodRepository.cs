using Azure.Core;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Infrastructure.Context;

namespace WriteBalance.Infrastructure.Repositories
{
    public class PeriodRepository : IPeriodRepository
    {
        private readonly AppDbContext _context;

        private readonly ILogger<PeriodRepository> _logger;

        public PeriodRepository(AppDbContext context, ILogger<PeriodRepository> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task<int> GetCompanyIdAsync(APIRequestDto request)
        {
            try
            {
                _logger.LogInformation("Starting GetCompanyIdAsync...");
                var entity = await _context.Periods
                    .AsNoTracking()
                    .FirstOrDefaultAsync(x => x.Id == request.PeriodId);

                return entity.CompanyId;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to fetch CompanyId from database (periodId): {request.PeriodId}");

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
