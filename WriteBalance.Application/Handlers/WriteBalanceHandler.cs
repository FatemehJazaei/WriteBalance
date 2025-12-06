using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WriteBalance.Application.DTOs;
using WriteBalance.Domain.Entities;
using WriteBalance.Application.Interfaces;
using WriteBalance.Application.Exceptions;
using System.IO;
using System.Data;
using DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json;
using WriteBalance.Common.Logging;


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
        private readonly IFileEncoder _fileEncoder;
        private readonly ICheckInput _checkInput;
        private readonly Logger _logger;

        public WriteBalanceHandler(
            IAuthService authService,
            IApiService apiService,
            IBalanceGenerator balanceGenerator,
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
            _checkInput = checkInput;
            _fileEncoder = fileEncoder;
        }

        public async Task<bool> HandleAsync(APIRequestDto request, DBRequestDto requestDB)
        {

            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting HandleAsync"), $"WriteBalanceHandler: HandleAsync--typeReport:Info");

                if (requestDB.TarazType == "-1")
                {
                    var UserInterBalanceName = request.BalanceName;
                    var resultHamrah = false;
                    var resultSama = false;
                    var resultKarbordi = false;
                    var resultRayan = false;
                    var resultPouya = false;
                    var errors = new List<string>();

                    try
                    {
                        requestDB.TarazType = "1";
                        request.BalanceName = UserInterBalanceName + " سما";
                         resultSama = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                    }
                    catch(ConnectionMessageException ex)
                    {
                        resultSama = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در سما : " + m ));
                    }
                    try
                    {
                        requestDB.TarazType = "4";
                        request.BalanceName = UserInterBalanceName + " همراه";
                        resultHamrah = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                    }
                    catch (ConnectionMessageException ex)
                    {
                        resultHamrah = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در همراه :" + m ));
                    }
                    try
                    {
                        requestDB.TarazType = "3";
                        request.BalanceName = UserInterBalanceName + " کاربردی";
                        resultKarbordi = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                    }
                    catch (ConnectionMessageException ex)
                    {
                        resultKarbordi = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در کاربردی :" + m ));
                    }
                    try 
                    {
                        requestDB.TarazType = "2";
                        request.BalanceName = UserInterBalanceName + " رایان";
                        resultRayan = await Handle_Rayan_Async(request, requestDB);
                    }
                    catch (ConnectionMessageException ex)
                    {
                        resultRayan = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در رایان :" + m ));
                    }
                    try
                    {
                        requestDB.TarazType = "5";
                        request.BalanceName = UserInterBalanceName + " پویا";
                        resultPouya = await Handle_Poya_Async(request, requestDB);
                    }
                    catch (ConnectionMessageException ex)
                    {
                        resultRayan = false;
                        errors.AddRange(ex.ConnectionMessage.Messages.Select(m => " خطا در رایان :" + m));
                    }
                    if (resultSama && resultHamrah && resultKarbordi && resultRayan && resultPouya) 
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject("All results is true!"), $"WriteBalanceHandler: HandleAsync--typeReport:Info");
                        return await Task.FromResult(true);
                    }
                    else
                    {
                        Logger.WriteEntry(JsonConvert.SerializeObject($"resultSama: {resultSama}, resultHamrah: {resultHamrah}, resultKarbordi: {resultKarbordi}, resultRayan: {resultRayan},  resultPouya: {resultPouya}"), $"WriteBalanceHandler: HandleAsync--typeReport:Error");
                        throw new ConnectionMessageException(new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = errors
                            },
                            request.FolderPath
                            );
                    }

                }
                else if (requestDB.TarazType == "1" || requestDB.TarazType == "3" || requestDB.TarazType == "4")
                {
                    return await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                }
                else if (requestDB.TarazType == "2") 
                {
                     return await Handle_Rayan_Async(request, requestDB); ;
                }
                else if (requestDB.TarazType == "5") {
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
                    request.FolderPath
                    );
                    
                }

            }
            catch (ConnectionMessageException ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Error in  HandleAsync.{ex.Message}"), $"WriteBalanceHandler: HandleAsync--typeReport:Error");
                throw;
            }
        }


        public async Task<bool> Handle_Rayan_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            var excelStream = await _excelExporter.CreateWorkbookAsync();
            Logger.WriteEntry(JsonConvert.SerializeObject($"CreateWorkbookAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

             (var CompanyId, DateTime startTime, DateTime endTime ) = await _periodRepository.GetTimeAsync(request);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput( requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecuteRayanSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteRayanSPList done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");


            excelStream = await _balanceGenerator.GenerateRayanTablesAsync(financialRecord, _excelExporter, requestDB);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateRayanTablesAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            await _excelExporter.SaveUploadAsync(excelStream, request.FolderPath, request.FileName);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SaveUploadAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");


            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(requestDB.FolderPath, requestDB.FileName);
                Logger.WriteEntry(JsonConvert.SerializeObject($"EncodeFileToBase64Async done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetAccessTokenAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                bool IsUnique = await _apiService.GetVerifyUniqueNameAsync(token, request);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetVerifyUniqueNameAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, request);
                Logger.WriteEntry(JsonConvert.SerializeObject($"PostFileAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

                return await Task.FromResult(PostApi);
            }
            else
            {
                return await Task.FromResult(true);
            }
        }

        public async Task<bool> Handle_Poya_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            var excelStream = await _excelExporter.CreateWorkbookAsync();
            Logger.WriteEntry(JsonConvert.SerializeObject($"CreateWorkbookAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            (var CompanyId, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecutePoyaSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecutePoyaSPList done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            await _balanceGenerator.GeneratePoyaTablesAsync(financialRecord, _excelExporter, requestDB);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GeneratePoyaTablesAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(requestDB.FolderPath, requestDB.FileName);
                Logger.WriteEntry(JsonConvert.SerializeObject($"EncodeFileToBase64Async done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetAccessTokenAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, request);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetVerifyUniqueNameAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, request);
                Logger.WriteEntry(JsonConvert.SerializeObject($"PostFileAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

                return await Task.FromResult(PostApi);
            }
            else
            {
                return await Task.FromResult(true);
            }
        }

        public async Task<bool> Handle_Hamrah_Karbordi_Sama_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            var excelStream = await _excelExporter.CreateWorkbookAsync();
            Logger.WriteEntry(JsonConvert.SerializeObject($"CreateWorkbookAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            (var CompanyId, DateTime startTime, DateTime endTime) = await _periodRepository.GetTimeAsync(request);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            (string startTimeStr, string endTimeStr) = _checkInput.CheckDateInput(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"CheckDateInput done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecuteSPList(request, requestDB, startTimeStr, endTimeStr);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            excelStream = await _balanceGenerator.GenerateTablesAsync(financialRecord, _excelExporter, requestDB);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            await _excelExporter.SaveUploadAsync(excelStream, request.FolderPath, requestDB.FileName);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SaveUploadAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");


            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(requestDB.FolderPath, requestDB.FileName);
                Logger.WriteEntry(JsonConvert.SerializeObject($"EncodeFileToBase64Async done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetAccessTokenAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, request);
                Logger.WriteEntry(JsonConvert.SerializeObject($"GetVerifyUniqueNameAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, request);
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
