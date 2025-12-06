namespace WriteBalance.Application.DTOs
{
    public class DBRequestDto
    {
        public string UserNameDB { get; set; }
        public string PtokenDB { get; set; }
        public string ObjecttokenDB { get; set; }
        public string FromDateDB { get; set; }
        public string ToDateDB { get; set; }
        public string TarazType { get; set; }
        public string TarazTypePouya { get; set; }
        public string AllOrHasMandeh { get; set; }
        public string OrginalClientAddressDB { get; set; }
        public string FromVoucherNum { get; set; }
        public string ToVoucherNum { get; set; }
        public string ExceptVoucherNum { get; set; }
        public string OnlyVoucherNum { get; set; }
        public string PrintOrReport { get; set; }
        public string FileName { get; set; }
        public string FileNameRial { get; set; }
        public string FileNameArzi { get; set; }
        public string FolderPath { get; set; }
    }
}
