using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public  class PoyaFinancialRecord
    {
        public string Kol_Code { get; set; }
        public string Kol_Title { get; set; }

        public string Moeen_Code { get; set; }
        public string Moeen_Title { get; set; }

        public string Arz_Code { get; set; }

        public decimal Mande_Bed { get; set; }
        public decimal Mande_Bes { get; set; }

    }
}
