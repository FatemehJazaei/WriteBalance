using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Application.DTOs
{
    public class APIRequestDto
    {
        public string UserNameAPI { get; set; } = string.Empty;
        public string PasswordAPI { get; set; } = string.Empty;
        public int PeriodId { get; set; }
        public string BaseUrl { get; set; } = string.Empty;
        public string BalanceName { get; set; } = string.Empty;
        public int Delay { get; set; }
        public string tarazNameLatin { get; set; } = string.Empty;

    }
}
