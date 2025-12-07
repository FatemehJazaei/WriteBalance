using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Application.DTOs
{
    public class APIRequestDto
    {
        public string UserNameAPI { get; set; } 
        public string PasswordAPI { get; set; } 
        public int PeriodId { get; set; }

        public string BaseUrl { get; set; }
        public string BalanceName { get; set; }
        public string FileName { get; set; }
        public string FileNameArzi { get; set; }
        public string FileNameRial { get; set; }
        public string FolderPath { get; set; }
        public int Delay { get; set; }

    }
}
