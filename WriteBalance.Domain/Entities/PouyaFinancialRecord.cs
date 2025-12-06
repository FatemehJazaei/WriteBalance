using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public  class PouyaFinancialRecord
    {
        public int Taraz_Date { get; set; }
        public int Code_shobeh { get; set; }
        public string Kol_Code_Markazi { get; set; }
        public string Kol_Title { get; set; }
        public string Hesab_Code { get; set; }
        public int Kol_Code { get; set; }
        public int Arz_Code { get; set; }
        public int Moeen_Code { get; set; }
        public int Moeen { get; set; }
        public string Tafzili { get; set; }
        public string Code_Arz_Abbr { get; set; }
        public string Sharh_Arz { get; set; }
        public decimal Mande_Bed_arzi { get; set; }
        public decimal Mande_Bes_arzi { get; set; }
        public decimal Mande_Bed_rial { get; set; }
        public decimal Mande_Bes_rial { get; set; }
        public decimal Gardersh_Bed_rial { get; set; }
        public decimal Gardersh_Bes_rial { get; set; }
        public decimal Gardersh_Bed_arzi { get; set; }
        public decimal Gardersh_Bes_arzi { get; set; }


    }
}
