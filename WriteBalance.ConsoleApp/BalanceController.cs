using DocumentFormat.OpenXml.InkML;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Handlers;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;
using WriteBalance.Infrastructure.Services;

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
                // ورودی ها چک میشود
                var InputValid = _checkInput.CheckUserInput(config);
                //کدهای حذفی چک میشود
                List<ExceptCode> ExceptCodes = new List<ExceptCode>();
                if (config["tarazType"] == "2")
                {
                    ExceptCodes = _checkInput.CheckExceptCode(config);
                }

                string folderName = config["of"];
                string path = config["op"];

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                if (!Directory.Exists($"{path}/{folderName}"))
                    Directory.CreateDirectory($"{path}/{folderName}");

                string folderPath = Path.Combine(path, folderName);

                Logger.WriteEntry(JsonConvert.SerializeObject($"OutputPath: {folderPath}"), $"BalanceController--typeReport:Debug");
                //چون از سه مدیریت  ارتباط استفاده میکنیم، متغیرهای تراز های دیگر در این قسمت مقداردهی میشود
                // sama , karbourdi, hamrah 
                if (config["tarazType"] == "1" || config["tarazType"] == "3" ||
                    config["tarazType"] == "4")
                {
                    config["FromVoucherNum"] = "";
                    config["ToVoucherNum"] = "";
                    config["ExceptVoucherNum"] = "";
                    config["OnlyVoucherNum"] = "";
                    config["tarazTypePouya"] = "";
                    config["ExceptCode"] = "";
                }
                if(config["tarazType"] == "2")
                {
                    //rayan
                    config["tarazTypePouya"] = "";
                }
                if (config["tarazType"] == "5")
                {
                    //pouya
                    config["ExceptCode"] = "";
                    config["FromVoucherNum"] = "";
                    config["ToVoucherNum"] = "";
                    config["ExceptVoucherNum"] = "";
                    config["OnlyVoucherNum"] = "";

                }

                // برای انتقال اطلاعات ورودی کاربر به دیگر لایه ها
                var request = new APIRequestDto
                {
                    UserNameAPI = config["UserNameAPI"],
                    PasswordAPI = config["PasswordAPI"],
                    PeriodId = int.Parse(config["pi"]),
                    BaseUrl = config["AddressAPI"],
                    BalanceName = config["BalanceName"],
                    Delay = int.Parse(config["UploadTimeSpanSeconds"]),
                };

                // برای انتقال اطلاعات ورودی کاربر به دیگر لایه ها
                var requestDB = new DBRequestDto
                {
                    UserNameDB = config["UserNameDB"],
                    PtokenDB = config["ptokenDB"],
                    ObjecttokenDB = config["objecttokenDB"],
                    OrginalClientAddressDB = config["OrginalClientAddressDB"],
                    TarazType = config["tarazType"],
                    TarazTypePouya =  config["tarazTypePouya"],
                    AllOrHasMandeh = config["AllOrHasMandeh"],
                    GardeshOrMandeh = config["GardeshOrMandeh"],
                    FromDateDB = config["FromDateDB"],
                    ToDateDB = config["ToDateDB"],
                    FromVoucherNum = config["FromVoucherNum"],
                    ToVoucherNum = config["ToVoucherNum"],
                    ExceptVoucherNum = config["ExceptVoucherNum"],
                    OnlyVoucherNum = config["OnlyVoucherNum"],
                    PrintOrReport = config["PrintOrReport"],
                    TarazKolOrTarazMoeen = config["TarazKolOrTarazMoeen"],
                    FolderPath = folderPath,
                    FileName = "",
                    ExceptCode = ExceptCodes,
                };

                // اطلاعات به  کلاس مدیریت عملیات ارسال میشود و فرایند استارت میشود
                var result = await _writeBalanceHandler.HandleAsync(request, requestDB);
                // اگر عملیات با موفقیت انجام شود، کد 0 را برمیگرداند 
                if (result)
                {
                    Environment.ExitCode = 0;
                }
                else // اگر عملیات به صورت غیر قابل پیش بینی شکست بخورد، کد 604 را برمیگرداند 
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("Unhandled exception occurred in BalanceController - 604"), $"BalanceController--typeReport:Error");
                    Environment.ExitCode = 604;
                }

            }
            catch (ConnectionMessageException ex) // در صورتی که ارور شناخته شده ای رخ دهد  متن ارور ثبت میشود و کد -1 برگردانده میشود
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Unhandled exception occurred in BalanceController : {ex.Message}"), $"BalanceController--typeReport:Debug");

                File.WriteAllText($"{ex.FolderPath}/Messages.txt", JsonConvert.SerializeObject(ex.ConnectionMessage));
                Environment.ExitCode = -1;
            }
           
        }

    }
}
