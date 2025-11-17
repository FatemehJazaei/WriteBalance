using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Application.Interfaces
{
    public interface IExcelExporter
    {
        XLWorkbook GetWorkbookReport();
        XLWorkbook GetWorkbookUpload();
        Task<MemoryStream> CreateWorkbookAsync();

        Task SaveReportAsync(MemoryStream stream, string path, string fileName);
        Task SaveUploadAsync(MemoryStream stream, string path, string fileName);
    }
}
