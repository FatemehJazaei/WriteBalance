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
        Task<bool> GetVerifyUniqueNameAsync(string token, string FolderPath, string BalanceName);
        Task<bool> PostFileAsync(string token, string file, string FileName, string BalanceName, string description, int currencyType, string FolderPath);

    }
}
