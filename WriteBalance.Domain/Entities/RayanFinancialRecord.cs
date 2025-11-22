using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public  class RayanFinancialRecord
    {
        public string Group_code { get; set; }
        public string Group_Title { get; set; }

        public string Kol_Code { get; set; }
        public string Kol_Title { get; set; }

        public string Moeen_Code { get; set; }
        public string Moeen_Title { get; set; }

        public string Tafsili_Code { get; set; }
        public string Tafsili_Title { get; set; }

        public string joze1_Code { get; set; }
        public string joze1_Title { get; set; }

        public string joze2_Code { get; set; }
        public string joze2_Title { get; set; }

        public string? Code_Markaz_Hazineh { get; set; }
        public string? Code_Vahed_Amaliyat { get; set; }
        public string? Name_Vahed_Amaliyat { get; set; }
        public string? Code_Parvandeh { get; set; }
        public string? Name_Parvandeh { get; set; }

        public double Mandeh_Aval_dore { get; set; }
        public double bedehkar { get; set; }
        public double bestankar { get; set; }
        public double Mande_Bed { get; set; }
        public double Mande_Bes { get; set; }

    }
}
