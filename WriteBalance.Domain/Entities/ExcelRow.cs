using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public class ExcelRow
    {
        public string Col1 { get; set; } = string.Empty;
        public string Col2 { get; set; } = string.Empty;
        public decimal Col3 { get; set; }
        public decimal Col4 { get; set; }
        public decimal? Col5 { get; set; }
        public decimal? Col6 { get; set; }
    }
}
