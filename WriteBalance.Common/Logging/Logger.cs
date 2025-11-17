using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Common.Logging
{
    public class Logger
    {

        public static void WriteEntry(string jsondata, string infoType)
        {

            var path = System.IO.Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location);
            File.AppendAllText(path + ".txt", "Date : " + "--" +
                                                           DateTime.Now + " " + infoType + " : " + jsondata + "--" + Environment.NewLine);

        }
    }
}
