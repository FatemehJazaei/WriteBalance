using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public class GLFinancialRecord
    {
        public int? Branch_ID { get; set; }
        public string? RBank_Code { get; set; }
        public string? RBank_Title { get; set; }
        public int? FinApplication_ID { get; set; }
        public string? FinApplication_Title { get; set; }
        public int? Motamam { get; set; }
        public decimal? Remain_First_Credit { get; set; }
        public decimal? Remain_First_Debit { get; set; }
        public decimal? Flow_Credit { get; set; }
        public decimal? Flow_Debit { get; set; }
        public decimal? Remain_Last_Credit { get; set; }
        public decimal? Remain_last_Debit { get; set; }
        public decimal? Account_Remain { get; set; }

    }
}
