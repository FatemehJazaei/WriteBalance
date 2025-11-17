using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
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
        private readonly Logger _logger;

        public WriteBalanceHandler(
            IAuthService authService,
            IApiService apiService,
            IBalanceGenerator balanceGenerator,
            IFinancialRepository financialRepository,
            IExcelExporter excelExporter,
            IPeriodRepository periodRepository,
            IFileEncoder fileEncoder, Logger logger)
        {
            _authService = authService;
            _financialRepository = financialRepository;
            _apiService = apiService;
            _excelExporter = excelExporter;
            _periodRepository = periodRepository;
            _balanceGenerator = balanceGenerator;
            _fileEncoder = fileEncoder;
        }

        public async Task<bool> HandleAsync(APIRequestDto request, DBRequestDto requestDB)
        {

            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Starting HandleAsync"), $"WriteBalanceHandler: HandleAsync--typeReport:Info");

                if (requestDB.TarazType == "-1")
                {

                    var TarazNumb =new string[] { "1", "3", "4" };
                    var UserInterBalanceName = request.BalanceName;

                    foreach (var num in TarazNumb)
                    {

                        switch (num)
                        {
                            case "1":
                                request.BalanceName = UserInterBalanceName + "_سما_";
                                break;
                            case "4":
                                request.BalanceName = UserInterBalanceName + "_همراه_";
                                break;
                            case "3":
                                request.BalanceName = UserInterBalanceName + "_کاربردی_";
                                break;
                        }

                        requestDB.TarazType = num;
                        var result = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                        if (!result)
                        {
                            await Task.FromResult(result);
                        }

                    }
                    request.BalanceName = UserInterBalanceName + "_رایان_";
                    var resultRayan = await Handle_Rayan_Async(request, requestDB);
                    if (!resultRayan)
                    {
                        return await Task.FromResult(resultRayan);
                    }
                    return await Task.FromResult(resultRayan);
                    // request.BalanceName = UserInterBalanceName + "_پویا_";
                    // var resultPoya = await Handle_Poya_Async(request, requestDB);
                    //return await Task.FromResult(resultPoya);

                }
                else if (requestDB.TarazType == "1" || requestDB.TarazType == "3" || requestDB.TarazType == "4")
                {
                    return await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                }
                else if (requestDB.TarazType == "2") {
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
                throw ;
            }
        }


        public async Task<bool> Handle_Rayan_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            var excelStream = await _excelExporter.CreateWorkbookAsync();
            Logger.WriteEntry(JsonConvert.SerializeObject($"CreateWorkbookAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

             (var CompanyId, DateTime startTime, DateTime endTime ) = await _periodRepository.GetTimeAsync(request);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GetTimeAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            var financialRecord = _financialRepository.ExecuteRayanSPList(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteRayanSPList done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");


            excelStream = await _balanceGenerator.GenerateRayanTablesAsync(financialRecord, _excelExporter, request.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateRayanTablesAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");

            await _excelExporter.SaveUploadAsync(excelStream, request.FolderPath, request.FileName);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SaveUploadAsync done."), $"WriteBalanceHandler: Handle_Rayan_Async--typeReport:Info");


            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(request.FolderPath, request.FileName);
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

            var financialRecord = _financialRepository.ExecutePoyaSPList(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecutePoyaSPList done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            excelStream = await _balanceGenerator.GeneratePoyaTablesAsync(financialRecord, _excelExporter, request.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GeneratePoyaTablesAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            await _excelExporter.SaveUploadAsync(excelStream, request.FolderPath, request.FileName);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SaveUploadAsync done."), $"WriteBalanceHandler: Handle_Poya_Async--typeReport:Info");

            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(request.FolderPath, request.FileName);
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

            var financialRecord = _financialRepository.ExecuteSPList(requestDB, startTime, endTime);
            Logger.WriteEntry(JsonConvert.SerializeObject($"ExecuteSPList done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            excelStream = await _balanceGenerator.GenerateTablesAsync(financialRecord, _excelExporter, request.FolderPath);
            Logger.WriteEntry(JsonConvert.SerializeObject($"GenerateTablesAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");

            await _excelExporter.SaveUploadAsync(excelStream, request.FolderPath, request.FileName);
            Logger.WriteEntry(JsonConvert.SerializeObject($"SaveUploadAsync done."), $"WriteBalanceHandler: Handle_Hamrah_Karbordi_Sama_Async--typeReport:Info");


            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(request.FolderPath, request.FileName);
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
