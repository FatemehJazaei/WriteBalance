using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public class CompareRows
    {
        public string Code { get; set; } = string.Empty;
        public string Titel { get; set; } = string.Empty;
        public decimal? MandehGL { get; set; }
        public decimal? MnadehAll { get; set; }
        public decimal? Ekhtelaf { get; set; }

    }
}
