using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Infrastructure.Config
{
    public class ApiSettings
    {
        public string PostIsUniqueUrl { get; set; } = string.Empty;
        public string PostBalanceSheetUrl { get; set; } = string.Empty;

        public int ControllerName { get; set; } = 20;
        public int RetryCount { get; set; } = 3;
        public int RetryDelaySeconds { get; set; } = 2;
    }
}
