using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces.Repository
{
    public interface IPooyaCodingRepository
    {
        Task<List<PooyaCoding>> GetPooyaCodingAsync(string FolderPath);
    }
}
