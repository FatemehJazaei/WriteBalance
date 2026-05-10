using WriteBalance.Domain.Entities;

namespace WriteBalance.Application.DTOs
{
    public class DBRequestDto
    {
        public string UserNameDB { get; set; } = string.Empty;
        public string PtokenDB { get; set; } = string.Empty;
        public string ObjecttokenDB { get; set; } = string.Empty;
        public string ObjecttokenGL { get; set; } = string.Empty;
        
        public string FromDateDB { get; set; } = string.Empty;
        public string ToDateDB { get; set; } = string.Empty;
        public string TarazType { get; set; } = string.Empty;
        public string TarazTypePouya { get; set; } = string.Empty;
        public string AllOrHasMandeh { get; set; } = string.Empty;
        public string GardeshOrMandeh { get; set; } = string.Empty;
        public string OrginalClientAddressDB { get; set; } = string.Empty;
        public string FromVoucherNum { get; set; } = string.Empty;
        public string ToVoucherNum { get; set; } = string.Empty;
        public string PrintOrReport { get; set; } = string.Empty;
        public string TarazKolOrTarazMoeen { get; set; } = string.Empty;       
        public string FileName { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public List<ExceptCode> ExceptCode { get; set; } = new List<ExceptCode>();
        public List<string> ExceptVoucherNum { get; set; } = new List<string>();
        public string BeforeClose { get; set; } = string.Empty;
    }
}
