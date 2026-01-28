using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Common.Logging;
using WriteBalance.Domain.Entities;


namespace WriteBalance.Application.Handlers
{
    public class WriteBalanceHandler
    {

        private readonly IAuthService _authService;
        private readonly IApiService _apiService;
        private readonly IFinancialRepository _financialRepository;
        private readonly IExcelExporter _excelExporter;
        private readonly IPeriodRepository _periodRepository;
        private readonly IBalanceGenerator _balanceGenerator;
        private readonly IPouyaBalanceGenerator _pouyaBalanceGenerator;
        private readonly IRayanBalanceGenerator _rayanBalanceGenerator;
        private readonly IFileEncoder _fileEncoder;
        private readonly ICheckInput _checkInput;
        private readonly Logger _logger;

        public WriteBalanceHandler(
            IAuthService authService,
            IApiService apiService,
            IBalanceGenerator balanceGenerator,
            IPouyaBalanceGenerator pouyaBalanceGenerator,
            IRayanBalanceGenerator rayanBalanceGenerator,
            IFinancialRepository financialRepository,
            IExcelExporter excelExporter,
            IPeriodRepository periodRepository,
            ICheckInput checkInput,
            IFileEncoder fileEncoder, Logger logger)
        {
            _authService = authService;
            _financialRepository = financialRepository;
            _apiService = apiService;
            _excelExporter = excelExporter;
            _periodRepository = periodRepository;
            _balanceGenerator = balanceGenerator;
            _pouyaBalanceGenerator = pouyaBalanceGenerator;
            _rayanBalanceGenerator = rayanBalanceGenerator;
            _checkInput = checkInput;
            _fileEncoder = fileEncoder;
            _logger = logger;
        }

        public async Task<bool> HandleAsync(APIRequestDto request, DBRequestDto requestDB)
        {

            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting HandleAsync"), $"WriteBalanceHandler: HandleAsync--typeReport:Info");

                // با توجه به نوع تراز عملیات انتخاب میشود 
                // سما، همراه و کاربردی 
                if (requestDB.TarazType == "-1")
                {
                    var resultHamrah = false;
                    var resultSama = false;
                    var resultKarbordi = false;
                    var errors = new List<string>();

                    try
                    {
                        //سما
                        requestDB.TarazType = "1";
                        resultSama = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                    }
                    catch (ConnectionMessageException ex)
                    {

                        resultSama = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در سما : " + m));
                    }
                    try
                    {
                        // همراه
                        requestDB.TarazType = "4";
                        resultHamrah = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                    }
                    catch (ConnectionMessageException ex)
                    {

                        resultHamrah = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در همراه :" + m));
                    }
                    try
                    {
                        //کاربردی
                        requestDB.TarazType = "3";
                        resultKarbordi = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                    }
                    catch (ConnectionMessageException ex)
                    {
                        resultKarbordi = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در کاربردی :" + m));
                    }
                
                    // هر سه تراز با موفقیت انجام شود
                    if (resultSama && resultHamrah && resultKarbordi)
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject("All results is true!"), $"WriteBalanceHandler: HandleAsync--typeReport:Info");
                        return await Task.FromResult(true);
                    }
                    else
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"resultSama: {resultSama}, resultHamrah: {resultHamrah}, resultKarbordi: {resultKarbordi}"), $"WriteBalanceHandler: HandleAsync--typeReport:Error");
                        throw new ConnectionMessageException(new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = errors
                        },
                            requestDB.FolderPath
                        );
                    }

                }
                else if (requestDB.TarazType == "1" || requestDB.TarazType == "3" || requestDB.TarazType == "4") // یکی از تراز های سما، همراه و کاربردی
                {
                    return await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                }
                else if (requestDB.TarazType == "2") // تراز رایان
                {
                    return await Handle_Rayan_Async(request, requestDB); ;
                }
                else if (requestDB.TarazType == "5") // تراز پویا
                {
                    return await Handle_Poya_Async(request, requestDB);
                }
                else
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("TarazType is not found!"), $"WriteBalanceHandler: HandleAsync--typeReport:Error");

                    throw new ConnectionMessageException(new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { "تراز شناسایی نشد" }
                    },
                    requestDB.FolderPath
                    );

                }

            }
            catch (ConnectionMessageException ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Error in  HandleAsync.{ex.Message}"), $"WriteBalanceHandler: HandleAsync--typeReport:Error");
                throw;
            }
        }

        // ساخت تراز رایان
        public async Task<bool> Handle_Rayan_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            //  تنظیم نام تراز + تاریخ
            var pc = new PersianCalendar();
            var now = DateTime.Now;

            string timestamp = $"{pc.GetSecond(now):00}_{pc.GetMinute(now):00}-{pc.GetHour(now):00}" +
                               $"_{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";

            var balanceName = request.BalanceName;
            string isGardesh = "";
            if (requestDB.GardeshOrMandeh == "2") { isGardesh = "گردش"; }
            request.BalanceName = $"{balanceName} تراز {isGardesh} رایان {timestamp}";

            (var CompanyId, bool isClosed, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request, requestDB.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecuteRayanSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteRayanSPList done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            var excelStream = new MemoryStream();
            if (requestDB.GardeshOrMandeh == "1")
            {
                excelStream = await _rayanBalanceGenerator.GenerateRayanTablesAsync(financialRecord, _excelExporter, requestDB);
                request.tarazNameLatin = $"{request.tarazNameLatin}_Mandeh";
            }
            else if(requestDB.GardeshOrMandeh == "2")
            {
                requestDB.FileName = requestDB.FileName.Replace("تراز", "تراز گردش");
                excelStream = await _rayanBalanceGenerator.GenerateRayanGardeshTablesAsync(financialRecord, _excelExporter, requestDB);
                request.tarazNameLatin = $"{request.tarazNameLatin}_Gardesh";
            }
                
            Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateRayanTablesAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            if (requestDB.PrintOrReport == "1")
            {
                if (isClosed)
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $".در دوره مالی بسته یا غیرفعال، بارگذاری تراز امکان پذیر نیست" }
                        },
                    requestDB.FolderPath
                    );
                }
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(excelStream, requestDB.FolderPath, requestDB.FileName);
                Logger.WriteEntry(JsonConvert.SerializeObject($"EncodeFileToBase64Async done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetAccessTokenAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, requestDB.FolderPath, request.BalanceName);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetVerifyUniqueNameAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                bool PostApi = await _apiService.PostFileAsync( token, fileBase64, requestDB.FileName, request.BalanceName, request.tarazNameLatin, 1, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"PostFileAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                return await Task.FromResult(PostApi);
            }
            else
            {
                return await Task.FromResult(true);
            }
        }

        // مدیریت تولید تراز پویا : ارزی و ریالی
        public async Task<bool> Handle_Poya_Async(APIRequestDto request, DBRequestDto requestDB)
        {

            (var CompanyId, bool isClosed, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request, requestDB.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecutePoyaSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecutePoyaSPList done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            var excelStreamRiali = new MemoryStream();
            var excelStreamArzi = new MemoryStream();
            if (requestDB.GardeshOrMandeh == "1")
            {
                (excelStreamRiali, excelStreamArzi) = await _pouyaBalanceGenerator.GeneratePoyaTablesAsync(financialRecord, _excelExporter, requestDB);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GeneratePoyaTablesAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");
                request.tarazNameLatin = $"{request.tarazNameLatin}_Mandeh";
            }
            else if (requestDB.GardeshOrMandeh == "2") 
            {
                requestDB.FileName = requestDB.FileName.Replace("تراز", "تراز گردش");
                (excelStreamRiali, excelStreamArzi) = await _pouyaBalanceGenerator.GeneratePoyaGardeshTablesAsync(financialRecord, _excelExporter, requestDB);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GeneratePoyaTablesAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");
                request.tarazNameLatin = $"{request.tarazNameLatin}_Gardesh";
            }


            if (requestDB.PrintOrReport == "1")
            {
                if (isClosed)
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $".در دوره مالی بسته یا غیرفعال، بارگذاری تراز امکان پذیر نیست" }
                        },
                    requestDB.FolderPath
                    );
                }
                var token = await _authService.GetAccessTokenAsync(request, CompanyId, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetAccessTokenAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");


                var pc = new PersianCalendar();
                var now = DateTime.Now;

                string timestamp = $"{pc.GetSecond(now):00}_{pc.GetMinute(now):00}-{pc.GetHour(now):00}" +
                                   $"_{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";
                var balanceName = request.BalanceName;
                string isGardesh = "";
                if (requestDB.GardeshOrMandeh == "2") { isGardesh = "گردش"; }
                string BalanceNameArzi = $"{balanceName} تراز {isGardesh} ارزی پویا {timestamp}";
                //Arzi
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(excelStreamArzi,requestDB.FolderPath, requestDB.FileName.Replace("تراز", "تراز ارزی"));
                Logger.WriteEntry(JsonConvert.SerializeObject($"EncodeFileToBase64Async done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, requestDB.FolderPath, BalanceNameArzi);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetVerifyUniqueNameAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                bool PostApiArzi = await _apiService.PostFileAsync(token, fileBase64, requestDB.FileName.Replace("تراز", "تراز ارزی"), BalanceNameArzi, $"{request.tarazNameLatin}_Arzi", 2, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"PostFileAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                //Delay
                int DeleyMiliSec = (request.Delay + 1) * 1000;
                await Task.Delay(DeleyMiliSec);
                Logger.WriteEntry(JsonConvert.SerializeObject($"DeleyMiliSec: {DeleyMiliSec}"), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                //Riali
                string BalanceNameRial = $"{balanceName} تراز {isGardesh} ریالی پویا {timestamp}";

                token = await _authService.GetAccessTokenAsync(request, CompanyId, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetAccessTokenAsync done again."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                fileBase64 = await _fileEncoder.EncodeFileToBase64Async(excelStreamRiali, requestDB.FolderPath, requestDB.FileName.Replace("تراز", "تراز ریالی"));
                Logger.WriteEntry(JsonConvert.SerializeObject($"EncodeFileToBase64Async done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, requestDB.FolderPath, BalanceNameRial);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetVerifyUniqueNameAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, requestDB.FileName.Replace("تراز", "تراز ریالی"), BalanceNameRial, $"{request.tarazNameLatin}_Riali", 1, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"PostFileAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");


                if (PostApiArzi && PostApi)
                    return await Task.FromResult(true);
                else
                {
                    return await Task.FromResult(false);
                }
            }
            else
            {
                return await Task.FromResult(true);
            }
        }

        // مدیریت تولید تراز سما و همراه و کاربردی
        public async Task<bool> Handle_Hamrah_Karbordi_Sama_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            var pc = new PersianCalendar();
            var now = DateTime.Now;

            string timestamp = $"{pc.GetSecond(now):00}_{pc.GetMinute(now):00}-{pc.GetHour(now):00}" +
                               $"_{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";
            var balanceName = request.BalanceName;
            string isGardesh = "";
            if (requestDB.GardeshOrMandeh == "2") { isGardesh = "گردش"; }
            switch (requestDB.TarazType)
            {
                case "1":
                    request.BalanceName = $"{balanceName} تراز {isGardesh} سما {timestamp}";
                    break;
                case "3":
                    request.BalanceName = $"{balanceName} تراز {isGardesh} کاربردی {timestamp}";
                    break;
                case "4":
                    request.BalanceName = $"{balanceName} تراز {isGardesh} همراه {timestamp}";
                    break;
            }

            (var CompanyId, bool isClosed, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request, requestDB.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            var excelStream = new MemoryStream();
            if (requestDB.GardeshOrMandeh == "1")
            {
                excelStream = await _balanceGenerator.GenerateTablesAsync(financialRecord, _excelExporter, requestDB);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");
                request.tarazNameLatin = $"{request.tarazNameLatin}_Mandeh";
            }
            else if (requestDB.GardeshOrMandeh == "2") 
            {
                requestDB.FileName = requestDB.FileName.Replace("تراز", "تراز گردش");
                excelStream = await _balanceGenerator.GenerateGardeshTablesAsync(financialRecord, _excelExporter, requestDB);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateGardeshTablesAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");
                request.tarazNameLatin = $"{request.tarazNameLatin}_Gardesh";
            }

            if (requestDB.PrintOrReport == "1")
            {
                if (isClosed)
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $".در دوره مالی بسته یا غیرفعال، بارگذاری تراز امکان پذیر نیست" }
                        },
                    requestDB.FolderPath
                    );
                }

                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(excelStream, requestDB.FolderPath, requestDB.FileName);
                Logger.WriteEntry(JsonConvert.SerializeObject($"EncodeFileToBase64Async done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetAccessTokenAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, requestDB.FolderPath, request.BalanceName);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetVerifyUniqueNameAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, requestDB.FileName, request.BalanceName, request.tarazNameLatin, 1, requestDB.FolderPath);
                Logger.WriteEntry(JsonConvert.SerializeObject($"PostFileAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

                return await Task.FromResult(PostApi);
            }
            else
            {
                return await Task.FromResult(true);
            }

        }
    }
}
