using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Application.DTOs
{
    public class ConnectionMessage
    {
        public MessageType MessageType { get; set; } = MessageType.Error;
        public List<string> Messages { get; set; } = new List<string>();
    }

    public enum MessageType
    {
        Error = 1,
        Warning = 2
    }

}
