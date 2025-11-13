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
        XLWorkbook GetWorkbook();
        Task<MemoryStream> CreateWorkbookAsync();

        Task SaveAsync(MemoryStream stream, string path, string fileName);
    }
}
