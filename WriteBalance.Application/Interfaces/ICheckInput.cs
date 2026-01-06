using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.Interfaces
{
    public interface ICheckInput
    {
        (string, string) CheckDateInput(DBRequestDto requestDB, DateTime startDateTime, DateTime endDateTime);
        bool CheckUserInput(Dictionary<string, string> config);
        List<ExceptCode> CheckExceptCode(Dictionary<string, string> config);
    }
}
