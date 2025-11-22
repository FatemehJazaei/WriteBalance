using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;

namespace WriteBalance.Application.Interfaces
{
    public interface ICheckInput
    {
        (string, string) CheckDateInput(DBRequestDto requestDB, DateTime startDateTime, DateTime endDateTime);
        bool CheckUserInput(Dictionary<string, string> config);
    }
}
