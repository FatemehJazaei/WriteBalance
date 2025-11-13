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
        private readonly ILogger<WriteBalanceHandler> _logger;

        public WriteBalanceHandler(
            IAuthService authService,
            IApiService apiService,
            IBalanceGenerator balanceGenerator,
            IFinancialRepository financialRepository,
            IExcelExporter excelExporter,
            IPeriodRepository periodRepository,
            IFileEncoder fileEncoder,
            ILogger<WriteBalanceHandler> logger)
        {
            _authService = authService;
            _financialRepository = financialRepository;
            _apiService = apiService;
            _excelExporter = excelExporter;
            _periodRepository = periodRepository;
            _balanceGenerator = balanceGenerator;
            _fileEncoder = fileEncoder;
            _logger = logger;
        }

        public async Task<bool> HandleAsync(APIRequestDto request, DBRequestDto requestDB)
        {

            try
            {
                _logger.LogInformation("Starting HandleAsync...");

                if (requestDB.TarazType == "0")
                {
                    var TarazNumb =new string[] { "1", "3", "4" };
                    foreach (var num in TarazNumb)
                    {
                        requestDB.TarazType = num;
                        var result = await Handle_Hamrah_Karbordi_Sama_Async(request, requestDB);
                        if (!result)
                        {
                            await Task.FromResult(result);
                        }

                    }
                    var resultRayan = await Handle_Rayan_Async(request, requestDB);
                    if (!resultRayan)
                    {
                        return await Task.FromResult(resultRayan);
                    }
                    var resultPoya = await Handle_Poya_Async(request, requestDB);
                    return await Task.FromResult(resultPoya);

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
                    _logger.LogError( "Error in  GenerateExcelHandler file.TarazType is not found!");
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
                _logger.LogError(ex, "Error in  GenerateExcelHandler file. HandleAsync fail!");
                throw ;
            }
        }


        public async Task<bool> Handle_Rayan_Async(APIRequestDto request, DBRequestDto requestDB)
        {
            var excelStream = await _excelExporter.CreateWorkbookAsync();
            _logger.LogInformation("CreateWorkbookAsync done Successfully.");

            var financialRecord = _financialRepository.ExecuteRayanSPList(requestDB);
            _logger.LogInformation("ExecuteMainProc done Successfully.");

            excelStream = await _balanceGenerator.GenerateRayanTablesAsync(financialRecord, _excelExporter, request.FolderPath);
            _logger.LogInformation("GenerateTablesAsync done Successfully.");

            await _excelExporter.SaveAsync(excelStream, request.FolderPath, request.FileName);
            _logger.LogInformation("SaveAsync done Successfully.");


            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(request.FolderPath, request.FileName);
                _logger.LogInformation("EncodeFileToBase64Async done Successfully.");

                var CompanyId = await _periodRepository.GetCompanyIdAsync(request);
                _logger.LogInformation("GetCompanyIdAsync done Successfully.");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId);
                _logger.LogInformation("GetAccessTokenAsync done Successfully.");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, request);

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, request);

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
            _logger.LogInformation("CreateWorkbookAsync done Successfully.");

            var financialRecord = _financialRepository.ExecutePoyaSPList(requestDB);
            _logger.LogInformation("ExecuteMainProc done Successfully.");

            excelStream = await _balanceGenerator.GeneratePoyaTablesAsync(financialRecord, _excelExporter, request.FolderPath);
            _logger.LogInformation("GenerateTablesAsync done Successfully.");

            await _excelExporter.SaveAsync(excelStream, request.FolderPath, request.FileName);
            _logger.LogInformation("SaveAsync done Successfully.");

            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(request.FolderPath, request.FileName);
                _logger.LogInformation("EncodeFileToBase64Async done Successfully.");

                var CompanyId = await _periodRepository.GetCompanyIdAsync(request);
                _logger.LogInformation("GetCompanyIdAsync done Successfully.");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId);
                _logger.LogInformation("GetAccessTokenAsync done Successfully.");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, request);

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, request);

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
            _logger.LogInformation("CreateWorkbookAsync done Successfully.");

            var financialRecord = _financialRepository.ExecuteSPList(requestDB);
            _logger.LogInformation("ExecuteMainProc done Successfully.");

            excelStream = await _balanceGenerator.GenerateTablesAsync(financialRecord, _excelExporter, request.FolderPath);
            _logger.LogInformation("GenerateTablesAsync done Successfully.");

            await _excelExporter.SaveAsync(excelStream, request.FolderPath, request.FileName);
            _logger.LogInformation("SaveAsync done Successfully.");


            if (requestDB.PrintOrReport == "1")
            {
                var fileBase64 = await _fileEncoder.EncodeFileToBase64Async(request.FolderPath, request.FileName);
                _logger.LogInformation("EncodeFileToBase64Async done Successfully.");

                var CompanyId = await _periodRepository.GetCompanyIdAsync(request);
                _logger.LogInformation("GetCompanyIdAsync done Successfully.");

                var token = await _authService.GetAccessTokenAsync(request, CompanyId);
                _logger.LogInformation("GetAccessTokenAsync done Successfully.");

                _ = await _apiService.GetVerifyUniqueNameAsync(token, request);

                bool PostApi = await _apiService.PostFileAsync(token, fileBase64, request);

                return await Task.FromResult(PostApi);
            }
            else
            {
                return await Task.FromResult(true);
            }

        }
    }
}
