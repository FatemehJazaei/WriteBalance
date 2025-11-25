using DocumentFormat.OpenXml.InkML;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Handlers;
using WriteBalance.Application.Exceptions;
using WriteBalance.Infrastructure.Services;
using WriteBalance.Common.Logging;
using WriteBalance.Application.Interfaces;

namespace WriteBalanceConsoleApp
{
    public class BalanceController
    {
        private readonly WriteBalanceHandler _writeBalanceHandler;
        private readonly ICheckInput _checkInput;
        public BalanceController( WriteBalanceHandler writeBalanceHandler, ICheckInput checkInput) 
        {
            _writeBalanceHandler = writeBalanceHandler; 
            _checkInput = checkInput;

        }
        public async Task InputBalanceController(Dictionary<string, string> config)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting InputBalanceController ..."), $"BalanceController--typeReport:Info");

                var InputValid = _checkInput.CheckUserInput(config);

                string folderName = config["of"];
                string path = config["op"];

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                if (!Directory.Exists($"{path}/{folderName}"))
                    Directory.CreateDirectory($"{path}/{folderName}");

                string folderPath = Path.Combine(path, folderName);

                Logger.WriteEntry(JsonConvert.SerializeObject($"OutputPath: {folderPath}"), $"BalanceController--typeReport:Debug");

                var request = new APIRequestDto
                {
                    UserNameAPI = config["UserNameAPI"],
                    PasswordAPI = config["PasswordAPI"],
                    PeriodId = int.Parse(config["pi"]),
                    BaseUrl = config["AddressAPI"],
                    BalanceName = config["BalanceName"],
                    FolderPath = folderPath,
                    FileName = "",
                };


                var requestDB = new DBRequestDto
                {
                    UserNameDB = config["UserNameDB"],
                    PtokenDB = config["ptokenDB"],
                    ObjecttokenDB = config["objecttokenDB"],
                    OrginalClientAddressDB = config["OrginalClientAddressDB"],
                    TarazType = config["tarazType"],
                    AllOrHasMandeh = config["AllOrHasMandeh"],
                    FromDateDB = config["FromDateDB"],
                    ToDateDB = config["ToDateDB"],
                    FromVoucherNum = config["FromVoucherNum"],
                    ToVoucherNum = config["ToVoucherNum"],
                    ExceptVoucherNum = config["ExceptVoucherNum"],
                    OnlyVoucherNum = config["OnlyVoucherNum"],
                    PrintOrReport = config["PrintOrReport"],
                    FolderPath = folderPath,
                    FileName = "",
                };

                var result = await _writeBalanceHandler.HandleAsync(request, requestDB);

                if (result)
                {
                    Environment.ExitCode = 0;
                }
                else 
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("Unhandled exception occurred in BalanceController - 604"), $"BalanceController--typeReport:Error");
                    Environment.ExitCode = 604;
                }

            }
            catch (ConnectionMessageException ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Unhandled exception occurred in BalanceController : {ex.Message}"), $"BalanceController--typeReport:Debug");

                File.WriteAllText($"{ex.FolderPath}/Messages.txt", JsonConvert.SerializeObject(ex.ConnectionMessage));
                Environment.ExitCode = -1;
            }
           
        }

    }
}
