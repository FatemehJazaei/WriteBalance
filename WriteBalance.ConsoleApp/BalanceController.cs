using DocumentFormat.OpenXml.InkML;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Handlers;
using WriteBalance.Application.Exceptions;
using WriteBalance.Infrastructure.Services;
using ILogger = Microsoft.Extensions.Logging.ILogger;

namespace WriteBalance.ConsoleApp
{
    public class BalanceController
    {
        private readonly ILogger _logger;
        private readonly WriteBalanceHandler _writeBalanceHandler;
        public BalanceController(ILogger<BalanceGenerator> logger, WriteBalanceHandler writeBalanceHandler) 
        {
            _writeBalanceHandler = writeBalanceHandler;

        }
        public async Task InputBalanceController(Dictionary<string, string> config)
        {
            try
            {
                string folderName = config["of"];
                string path = config["op"];

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                if (!Directory.Exists($"{path}/{folderName}"))
                    Directory.CreateDirectory($"{path}/{folderName}");

                var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string fileName = $"{config["BalanceName"]}_{timestamp}.xlsx";
                string folderPath = Path.Combine(path, folderName);

                Log.Information($"OutputPath: {folderPath}");

                var request = new APIRequestDto
                {
                    UserNameAPI = config["UserNameAPI"],
                    PasswordAPI = config["PasswordAPI"],
                    PeriodId = int.Parse(config["pi"]),
                    BaseUrl = config["AddressAPI"],
                    BalanceName = config["BalanceName"],
                    FolderPath = folderPath,
                    FileName = fileName,
                };

                var requestDB = new DBRequestDto
                {
                    UserNameDB = config["UserNameDB"],
                    PtokenDB = config["ptokenDB"],
                    ObjecttokenDB = config["objecttokenDB"],
                    OrginalClientAddressDB = config["OrginalClientAddressDB"],
                    TarazType = config["OrginalClientAddressDB"],
                    FromDateDB = config["OrginalClientAddressDB"],
                    ToDate = config["OrginalClientAddressDB"],
                    FromVoucherNum = config["OrginalClientAddressDB"],
                    ToVoucherNum = config["OrginalClientAddressDB"],
                    ExceptVoucherNum = config["1"],
                    OnlyVoucherNum = config["1"],
                    PrintOrReport = config["1"],
                    FolderPath = folderPath,
                    FileName = fileName,
                };

                var result = await _writeBalanceHandler.HandleAsync(request, requestDB);

                if (result)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Environment.ExitCode = 0;
                }
                else 
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Environment.ExitCode = 604;
                }

            }
            catch (ConnectionMessageException ex)
            {
                Log.Error(ex.Message);
                Console.ForegroundColor = ConsoleColor.Red;

                File.WriteAllText($"{ex.FolderPath}/Messages.txt", JsonConvert.SerializeObject(ex.ConnectionMessage));
                Environment.ExitCode = -1;
            }
           
        }

    }
}
