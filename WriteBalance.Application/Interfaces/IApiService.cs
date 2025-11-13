using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;

namespace WriteBalance.Application.Interfaces
{
    public interface IApiService
    {
        Task<bool> GetVerifyUniqueNameAsync(string token, APIRequestDto request);
        Task<bool> PostFileAsync(string token, string name, APIRequestDto request);
    }
}
