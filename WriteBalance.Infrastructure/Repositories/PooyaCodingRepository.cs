using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Application.Interfaces.Repository;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;
using WriteBalance.Infrastructure.Context;

namespace WriteBalance.Infrastructure.Repositories
{
     public class PooyaCodingRepository : IPooyaCodingRepository
    {
        private readonly ModulesDbContext _context;

        public PooyaCodingRepository(ModulesDbContext context)
        {
            _context = context;
        }

        public List<PooyaCoding> GetPooyaCodingAsync(string FolderPath)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting GetPooyaCodingAsync"), $"PooyaCodingRepository:PooyaCodingRepository--typeReport:Info");

                var entities = _context.PooyaCodings
                    .AsNoTracking()
                    .ToList();

                Logger.WriteEntry(JsonConvert.SerializeObject($"entities.Count : {entities.Count}"), $"PooyaCodingRepository:PooyaCodingRepository--typeReport:Info");
                return entities;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"PooyaCodingRepository:GetPooyaCodingAsync--typeReport:Error");
              
                throw new ConnectionMessageException(new ConnectionMessage
                {
                    MessageType = MessageType.Error,
                    Messages = new List<string> { "ارتباط با پایگاه داده اکسیر ناموفق!" }
                },
                FolderPath
                );
            }
        }
    }
}
