using Azure.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces.Repository;
using WriteBalance.Domain.Entities;
using WriteBalance.Infrastructure.Context;

namespace WriteBalance.Infrastructure.Services
{
    public class GetPooyaCoding
    {
        private readonly ModulesDbContext _context;
        private readonly IPooyaCodingRepository pooyaCodingRepository;

        public GetPooyaCoding(ModulesDbContext context, IPooyaCodingRepository pooyaCodingRepository)
        {
            _context = context;
            pooyaCodingRepository = pooyaCodingRepository;
        }

        public async Task<Dictionary<string, string>> ExecuteAsync(string FolderPath)
        {
            try
            {
                var Coding = await pooyaCodingRepository.GetPooyaCodingAsync(FolderPath);
                var DicCode = CreateDicFromnCode(Coding, FolderPath);
                return DicCode;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public Dictionary<string, string> CreateDicFromnCode(List<PooyaCoding> pooyaCodings, string FolderPath)
        {
            try
            {
                var dictionary = pooyaCodings.ToDictionary(
                    x => $"{x.CodeKol}_{x.CodeArz}_{x.GroupMoein}",
                    x => $"{x.CodeKol}_{x.CodeOmoorMali:D4}"
                );

                return dictionary;
            }
            catch (Exception ex) 
            {
                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $" خطا در کدهای معادل پویا" }
                        },
                    FolderPath
                    );
            }
           
        }
    }
}
