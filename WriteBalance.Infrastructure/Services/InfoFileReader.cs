using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Infrastructure.Services
{
    public static class InfoFileReader
    {
        public static async Task<Dictionary<string, string>> ReadAsync(string[] args)
        {
            var config = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            config["op"] = "E:\\Projects";
            config["of"] = "WriteBalance";
            config["pi"] = "5046";
            config["AddressServerBank"] = "Exir-203";
            config["DataBaseNameBank"] = "Database1";
            config["UserNameDB"] = "SysSouratMali";
            config["ptokenDB"] = "c3d8e6a3459b15c9";
            config["objecttokenDB"] = "3d9758851923e42b";
            config["FromDateDB"] = "14030101";
            config["ToDate"] = "14031230";
            config["tarazType"] = "1";
            config["OrginalClientAddressDB"] = "10.15.52.97";

            config["BalanceName"] = "Balance";


            for (int i = 0; i < args.Length - 1; i += 2)
            {
                string key = args[i].TrimStart('-');
                string value = args[i + 1];
                config[key] = value;
            }

            string filePath = Path.Combine(AppContext.BaseDirectory, @"..\Basic_Information\Info.txt");
            filePath = Path.GetFullPath(filePath);

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Info.txt not found at: {filePath}");


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
    }
}
