using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Domain.Entities
{
    public class Period
    {
        public int Id { get; set; }
        public int CompanyId { get; set; }
        // دوره مالی غیر فعال 1
        // دوره مالی فعال  0
        public int StateType { get; set; }
        //دوره مالی بسته شده true
        //دوره مالی باز false
        public bool Closed  { get; set; }

        public DateTime StartDate { get; set; }
        public DateTime TimeEnd { get; set; }
    }
}
