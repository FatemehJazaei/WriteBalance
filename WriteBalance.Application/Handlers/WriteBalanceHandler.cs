using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
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
        private readonly IGLBalanceGenerator _gLBalanceGenerator;
        private readonly IAllBalanceGenerator _allBalanceGenerator;
        private readonly ICompareBalance _compareBalance;
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
            IGLBalanceGenerator gLBalanceGenerator,
            IAllBalanceGenerator allBalanceGenerator,
            IExcelExporter excelExporter,
            IPeriodRepository periodRepository,
            ICompareBalance compareBalance,
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
            _gLBalanceGenerator = gLBalanceGenerator;
            _allBalanceGenerator = allBalanceGenerator;
            _compareBalance = compareBalance;
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
                else if (requestDB.TarazType == "1" || requestDB.TarazType == "3" || requestDB.TarazType == "4" || requestDB.TarazType == "10" || requestDB.TarazType == "9") // یکی از تراز های سما، همراه و کاربردی
                {
                    return await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                }
                else if (requestDB.TarazType == "5") // تراز رایان
                {
                    return await Handle_Rayan_Async(request, requestDB); ;
                }
                else if (requestDB.TarazType == "2") // تراز پویا
                {
                    return await Handle_Poya_Async(request, requestDB);
                }
                else if (requestDB.TarazType == "6") // تراز GL
                {
                    return await Handle_GL_Async(request, requestDB);
                }
                else if (requestDB.TarazType == "7") // مقایسه تراز  جی ال  و پنج تراز دیگر 
                {
                    return await Handle_Compare_GL_Async(request, requestDB);
                }
                else if (requestDB.TarazType == "8") //  پنج تراز  
                {
                    return await Handle_5Balance_Async(request, requestDB);
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

            var financialRecord = _financialRepository.ExecuteRayanSPList(request, requestDB, startTimeStr, endTimeStr, false);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteRayanSPList done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            if (requestDB.ExceptVoucherNum.Count != 0)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"ExceptVoucherNum.Count: {requestDB.ExceptVoucherNum.Count} ."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");
                List<RayanFinancialRecord> ExceptRayanFinancialRecords = new List<RayanFinancialRecord>();
                foreach ( var VoucherNum in requestDB.ExceptVoucherNum)
                {
                    requestDB.ToVoucherNum = VoucherNum;
                    requestDB.FromVoucherNum = VoucherNum;
                    var exceptVouchernum = _financialRepository.ExecuteRayanSPList(request, requestDB, startTimeStr, endTimeStr,true);
                    Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteRayanSPList done for voucher number: {VoucherNum} ."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");
                    if (exceptVouchernum.Count != 0) 
                    {
                        ExceptRayanFinancialRecords.AddRange(exceptVouchernum);
                    }
         
                }
                ExceptRayanFinancialRecords = _rayanBalanceGenerator.ExceptRayanTables(ExceptRayanFinancialRecords, requestDB);
                Logger.WriteEntry(JsonConvert.SerializeObject($"ExceptRayanTables done ."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");
                financialRecord = financialRecord
                .Concat(ExceptRayanFinancialRecords)
                .ToList();
            }


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

            requestDB = _checkInput.CheckPouyaType(requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckPouyaType done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

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
                case "10":
                    //rayan
                    requestDB.TarazType = "5";
                    request.BalanceName = $"{balanceName} تراز {isGardesh} رایان {timestamp}";
                    break;
                case "9":
                    // pouya
                    requestDB.TarazType = "2";
                    request.BalanceName = $"{balanceName} تراز {isGardesh} پویا {timestamp}";
                    break;
            }

            (var CompanyId, bool isClosed, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request, requestDB.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            requestDB = _checkInput.CheckBeforeClose(requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckBeforeClose done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");


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


        // مدیریت تولید تراز GL
        public async Task<bool> Handle_GL_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            var pc = new PersianCalendar();
            var now = DateTime.Now;

            string timestamp = $"{pc.GetSecond(now):00}_{pc.GetMinute(now):00}-{pc.GetHour(now):00}" +
                               $"_{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";
            var balanceName = request.BalanceName;
            string isGardesh = "";
            if (requestDB.GardeshOrMandeh == "2") { isGardesh = "گردش"; }
            request.BalanceName = $"{balanceName} تراز {isGardesh} جی ال {timestamp}";

            (var CompanyId, bool isClosed, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request, requestDB.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_GL_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_GL_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecuteGLList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList done."), $"WriteBalanceHandler: Handle_GL_Async--typeReport:Info");

            var excelStream = new MemoryStream();
            if (requestDB.GardeshOrMandeh == "1")
            {
                excelStream = await _gLBalanceGenerator.GenerateGLTablesAsync(financialRecord, _excelExporter, requestDB);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync done."), $"WriteBalanceHandler: Handle_GL_Async--typeReport:Info");
                request.tarazNameLatin = $"{request.tarazNameLatin}_Mandeh";
            }
            else if (requestDB.GardeshOrMandeh == "2")
            {
                requestDB.FileName = requestDB.FileName.Replace("تراز", "تراز گردش");
                excelStream = await _gLBalanceGenerator.GenerateGardeshGLTablesAsync(financialRecord, _excelExporter, requestDB);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateGardeshTablesAsync done."), $"WriteBalanceHandler: Handle_GL_Async--typeReport:Info");
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


        // مقایسه تراز جی ال و تراز های دیگر 
        public async Task<bool> Handle_Compare_GL_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            
            var pc = new PersianCalendar();
            var now = DateTime.Now;

            string timestamp = $"{pc.GetSecond(now):00}_{pc.GetMinute(now):00}-{pc.GetHour(now):00}" +
                               $"_{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";

            (var CompanyId, bool isClosed, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request, requestDB.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            // تراز جی ال   
            var GLFinancialRecord = _financialRepository.ExecuteGLList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            var GLExcelRows = await _compareBalance.SetGLExcelRowAsync(GLFinancialRecord);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SetGLExcelRowAsync done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            // تراز سما  
            requestDB.TarazType = "1";
            var samaFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_sama done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            var samaExcelRow = await _compareBalance.SetExcelRowAsync(samaFinancialRecord);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SetExcelRowAsync_sama done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            // تراز همراه 
            requestDB.TarazType = "4";
            var hamrahFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_hamrah done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            var hamrahExcelRow = await _compareBalance.SetExcelRowAsync(hamrahFinancialRecord);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SetExcelRowAsync_hamrah done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            //تراز کاربردی 
            requestDB.TarazType = "3";
            var karbourdiFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_karbourdi done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            var karbourdiExcelRow = await _compareBalance.SetExcelRowAsync(karbourdiFinancialRecord);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SetExcelRowAsync_karbourdi done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            //تراز رایان 
            requestDB.TarazType = "5";
            var rayanFinancialRecord = _financialRepository.ExecuteRayanSPList(request, requestDB, startTimeStr, endTimeStr, false);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteRayanSPList done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            var rayanExcelRow = await _compareBalance.SetRayanExcelRowAsync(rayanFinancialRecord);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SetExcelRowAsync_rayan done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            //تراز پویا 
            requestDB.TarazType = "2";
            var pouyaFinancialRecord = _financialRepository.ExecutePoyaSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecutePoyaSPList done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            var pouyaExcelRow = await _compareBalance.SetPouyaExcelRowAsync(pouyaFinancialRecord);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SetPouyaExcelRowAsync done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            // تجمیع 5 تراز
            var allExcelRows = await _compareBalance.CreateAllExcelRowAsync(samaExcelRow, hamrahExcelRow, karbourdiExcelRow, rayanExcelRow, pouyaExcelRow);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CreateAllExcelRowAsync done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            // مقایسه تراز ها 
            var compareExcelRows = await _compareBalance.CompareBalanceAsync(allExcelRows, GLExcelRows, requestDB);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CompareBalanceAsync done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            // ذخیره در اکسل
            requestDB.FileName = $" گزارش مقایسه ترازها با تراز جی ال {timestamp}.xlsx";
            await _compareBalance.WriteExcelAsync(compareExcelRows, _excelExporter, requestDB);
            Logger.WriteEntry(JsonConvert.SerializeObject($"WriteExcelAsync done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            return await Task.FromResult(true);
        }

        // دریافت 5 تراز با هم دیگر به عنوان GL
        public async Task<bool> Handle_5Balance_Async(APIRequestDto request, DBRequestDto requestDB)
        {

            var pc = new PersianCalendar();
            var now = DateTime.Now;

            string timestamp = $"{pc.GetSecond(now):00}_{pc.GetMinute(now):00}-{pc.GetHour(now):00}" +
                               $"_{pc.GetDayOfMonth(now):00}_{pc.GetMonth(now):00}_{pc.GetYear(now):0000}";
            var balanceName = request.BalanceName;
            string isGardesh = "";
            if (requestDB.GardeshOrMandeh == "2") { isGardesh = "گردش"; }
            request.BalanceName = $"{balanceName} تراز جی ال {isGardesh} _ {timestamp}";


            (var CompanyId, bool isClosed, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request, requestDB.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            requestDB = _checkInput.CheckBeforeClose(requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckBeforeClose done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

            var errors = new List<string>();
            var resultHamrah = false;
            var resultSama = false;
            var resultKarbordi = false;
            var resultPouya = false;
            var resultRayan = false;
            var resultGl = false;

            List<ExcelRow> samaExcelRow = new List<ExcelRow>();
            List<ExcelRow> hamrahExcelRow = new List<ExcelRow>();
            List<ExcelRow> karbourdiExcelRow = new List<ExcelRow>();
            List<ExcelRow> rayanExcelRow = new List<ExcelRow>();
            List<ExcelRow> pouyaExcelRow = new List<ExcelRow>();

            List<FinancialRecord> samaFinancialRecord = new List<FinancialRecord>();
            List<FinancialRecord> hamrahFinancialRecord = new List<FinancialRecord>();
            List<FinancialRecord> karbourdiFinancialRecord = new List<FinancialRecord>();
            List<FinancialRecord> rayanFinancialRecord = new List<FinancialRecord>();
            List<FinancialRecord> pouyaFinancialRecord = new List<FinancialRecord>();
            try
            {
                // تراز سما  
                requestDB.TarazType = "1";
                 samaFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
                Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_sama done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

                samaExcelRow = await _allBalanceGenerator.CheckExcelRowGLAsync(samaFinancialRecord, requestDB);
                resultSama = true;
            }
            catch (ConnectionMessageException ex)
            {

                resultSama = false;
                errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در سما : " + m));
            }

            try
            {
                // تراز همراه 
                requestDB.TarazType = "4";
                hamrahFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
                Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_hamrah done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

                hamrahExcelRow = await _allBalanceGenerator.CheckExcelRowGLAsync(hamrahFinancialRecord, requestDB);
                resultHamrah= true;
            }
            catch (ConnectionMessageException ex)
            {

                resultHamrah = false;
                errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در همراه : " + m));
            }

            try
            {
                //تراز کاربردی 
                requestDB.TarazType = "3";
                karbourdiFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
                Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_karbourdi done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

                karbourdiExcelRow = await _allBalanceGenerator.CheckExcelRowGLAsync(karbourdiFinancialRecord, requestDB);
                resultKarbordi = true;
            }
            catch (ConnectionMessageException ex)
            {

                resultKarbordi = false;
                errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در کاربردی : " + m));
            }

            try
            {
                //تراز رایان 
                requestDB.TarazType = "5";
                rayanFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
                Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_rayan done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

                rayanExcelRow = await _allBalanceGenerator.CheckExcelRowGLAsync(rayanFinancialRecord, requestDB);
                resultRayan = true;
            }
            catch (ConnectionMessageException ex)
            {

                resultRayan = false;
                errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در رایان : " + m));
            }

            try
            {
                //تراز پویا 
                requestDB.TarazType = "2";
                pouyaFinancialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
                Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList_Poya  done."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");

                pouyaExcelRow = await _allBalanceGenerator.CheckExcelRowGLAsync(pouyaFinancialRecord, requestDB);
                resultPouya = true;
            }
            catch (ConnectionMessageException ex)
            {

                resultPouya = false;
                errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در پویا : " + m));
            }

            // تجمیع 5 تراز
            var AllExcelRow = samaExcelRow
                .Concat(hamrahExcelRow)
                .Concat(karbourdiExcelRow)
                .Concat(rayanExcelRow)
                .Concat(pouyaExcelRow)
                .ToList();

            var financialRecords = samaFinancialRecord
                .Concat(hamrahFinancialRecord)
                .Concat(karbourdiFinancialRecord)
                .Concat(rayanFinancialRecord)
                .Concat(pouyaFinancialRecord)
                .ToList();

            Logger.WriteEntry(JsonConvert.SerializeObject($"Create All Excel Row."), $"WriteBalanceHandler: Handle_Compare_GL_Async--typeReport:Info");


            try
            {
                requestDB.FileName = $" تراز جی ال دریافت شده در تاریخ  {timestamp} .xlsx";
                request.tarazNameLatin = "GL";

                var excelStream = new MemoryStream();
                if (requestDB.GardeshOrMandeh == "1")
                {
                    excelStream = await _allBalanceGenerator.GenerateAllTableAsync( financialRecords, AllExcelRow, _excelExporter, requestDB);
                    Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");
                    request.tarazNameLatin = $"{request.tarazNameLatin}_Mandeh";
                }
                else if (requestDB.GardeshOrMandeh == "2")
                {
                    requestDB.FileName = requestDB.FileName.Replace("تراز", "تراز گردش");
                    excelStream = await _allBalanceGenerator.GenerateAllTableGardeshAsync(financialRecords, AllExcelRow, _excelExporter, requestDB);
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

                    resultGl = PostApi;
                }
                else 
                {
                    resultGl = true;
                }
            }
            catch (ConnectionMessageException ex)
            {
                errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در جی ال : " + m));
            }


            // هر 6 تراز با موفقیت انجام شود
            if (resultSama && resultHamrah && resultKarbordi && resultRayan && resultPouya && resultGl)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("All results is true!"), $"WriteBalanceHandler: HandleAsync--typeReport:Info");
                return await Task.FromResult(true);
            }
            else
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"resultSama: {resultSama}, resultHamrah: {resultHamrah}, resultKarbordi: {resultKarbordi}, resultRayan: {resultRayan}, resultPouya:{resultPouya}, resultGl: {resultGl} "), $"WriteBalanceHandler: HandleAsync--typeReport:Error");
                throw new ConnectionMessageException(new ConnectionMessage
                {
                    MessageType = MessageType.Error,
                    Messages = errors
                },
                    requestDB.FolderPath
                );
            }

        }

    }
}
