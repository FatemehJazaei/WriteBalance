using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;

namespace WriteBalance.Application.Exceptions
{
    // برای ارسال خطا به اکسیر 
    public class ConnectionMessageException : Exception
    {
        public ConnectionMessage ConnectionMessage { get; }
        public string FolderPath { get; }
        public ConnectionMessageException(ConnectionMessage connectionMessage, string folderPath)
        {
            ConnectionMessage = connectionMessage;
            FolderPath = folderPath;

        }
    }
}
