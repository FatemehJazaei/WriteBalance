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
        XLWorkbook GetWorkbookGardesh();
        XLWorkbook GetWorkbookUploadArzi();
        XLWorkbook GetWorkbookPouyaNotFound();
        
        Task SaveReportAsync(MemoryStream stream, string path, string fileName);
        Task SaveUploadAsync(MemoryStream stream, string path, string fileName);
        Task SaveGardeshAsync(MemoryStream stream, string path, string fileName);
        Task SaveUploadArziAsync(MemoryStream stream, string path, string fileName);
        Task SavePouyaNotFoundAsync(MemoryStream stream, string path, string fileName);
    }
}
