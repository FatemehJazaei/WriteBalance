using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Application.Interfaces
{
    public interface IFileEncoder
    {
        Task<string> EncodeFileToBase64Async(string filePath, string fileName);
    }
}
