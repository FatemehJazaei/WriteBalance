using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public  class FinancialRecord
    {
        public string Kol_Code { get; set; }
        public string Kol_Title { get; set; }

        public string Moeen_Code { get; set; }
        public string Moeen_Title { get; set; }


        public long Mande_Bed { get; set; }
        public long Mande_Bes { get; set; }

    }
}
