using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Common.Logging;

namespace WriteBalance.Infrastructure.Services
{
    public static class InfoFileReader
    {
        public static async Task<Dictionary<string, string>> ReadAsync(string[] args)
        {
            try
            {

                var config = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);


                foreach (var arg in args)
                {
                    Logger.WriteEntry(arg, "InputArg");
                    var parts = arg.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);

                    if (parts.Length >= 2)
                    {
                        var value = "";
                        var key = parts[0].TrimStart('-');
                        for (var i = 1; i < parts.Length; i++)
                        {
                             value += parts[i];
                        }

                        config[key] = value;
                    }
                    if (parts.Length == 1)
                    {
                        var key = parts[0].TrimStart('-');
                        config[key] = "";
                    }

                }

                foreach (var kv in config)
                {

                    Logger.WriteEntry(JsonConvert.SerializeObject($"{kv.Key} = {kv.Value}"), $"InfoFileReader: HandleAsync--typeReport:Info");
                }

                config["UserNameDB"] = "SysSouratMali";
                config["ptokenDB"] = "c3d8e6a3459b15c9";
                config["objecttokenDB"] = "3d9758851923e42b";
                config["OrginalClientAddressDB"] = "10.15.52.97";
                config["FromVoucherNum"] = "";
                config["ToVoucherNum"] = "";
                config["ExceptVoucherNum"] = "";
                config["OnlyVoucherNum"] = "";


                config["AddressServerBank"] = "Exir-203";
                config["DataBaseNameBank"] = "Refah";
                config["op"] = "E:\\Projects";
                config["of"] = "WriteBalance";
                config["pi"] = "5046";
                config["tarazType"] = "5";
                config["tarazTypePouya"] = "1";
                config["AllOrHasMandeh"] = "1";
                config["PrintOrReport"] = "1";
                config["BalanceName"] = "BalanceTest8";
                config["FromDateDB"] = "14040104";
                config["ToDateDB"] = "14040424";

                string filePath = Path.Combine(AppContext.BaseDirectory, @"..\Basic_Information\Info.txt");
                filePath = Path.GetFullPath(filePath);

                if (!File.Exists(filePath))
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Info.txt not found at: {filePath}"), $"InfoFileReader: HandleAsync--typeReport:Error");
                    throw new FileNotFoundException($"Info.txt not found at: {filePath}");
                }


                var lines = await File.ReadAllLinesAsync(filePath);

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split('|', 2);
                    if (parts.Length == 2)
                        config[parts[0].Trim()] = parts[1].Trim();
                }

                return config;
            }
            catch(Exception ex) 
            {
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"InfoFileReader: HandleAsync--typeReport:Error");
                throw;

            }

            
        }
    }
}
