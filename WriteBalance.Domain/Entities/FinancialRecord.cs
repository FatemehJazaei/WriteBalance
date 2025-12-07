using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public  class FinancialRecord
    {
        public string? Kol_Code { get; set; }
        public string? Kol_Title { get; set; }
        public string? Moeen_Code { get; set; }
        public string? Moeen_Title { get; set; }
        public int? Tafzil_Code { get; set; }
        public int? Tafzil_Tilte { get; set; }
        public string? FinApplication_Title { get; set; }
        public int? AccountNature_ID { get; set; }
        public string? AccountNature_Title { get; set; }
        public int? Motamam { get; set; }
        public decimal? Remain_First_Credit { get; set; }
        public decimal? Remain_First_Debit { get; set; }
        public decimal? Flow_Credit { get; set; }
        public decimal? Flow_Debit { get; set; }
        public decimal? Remain_Last_Credit { get; set; }
        public decimal? Remain_last_Debit { get; set; }


        /*
        public string Kol_Code { get; set; }
        public string Kol_Title { get; set; }

        public string Moeen_Code { get; set; }
        public string Moeen_Title { get; set; }

        public Int32 Tafzil_Code { get; set; }
        public Int32 Tafzil_Tilte { get; set; }

        public string FinApplication_Title { get; set; }

        public decimal Gardersh_Bed { get; set; }
        public decimal Gardersh_Bes { get; set; }

        public decimal Mande_Bed { get; set; }
        public decimal Mande_Bes { get; set; }
        */
    }
}
