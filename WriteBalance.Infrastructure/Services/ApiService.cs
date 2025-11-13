using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Exceptions;
using WriteBalance.Application.Interfaces;
using WriteBalance.Infrastructure.Config;

namespace WriteBalance.Infrastructure.Services
{
    public class ApiService : IApiService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiConfig _settings;
        private readonly ILogger<ApiService> _logger;

        public ApiService(HttpClient httpClient, ApiConfig settings, ILogger<ApiService> logger)
        {
            _httpClient = httpClient;
            _settings = settings;
            _logger = logger;
        }

        public async Task<bool> GetVerifyUniqueNameAsync(string token, APIRequestDto request)
        {
            try
            {
                _logger.LogInformation($"Starting GetVerifyUniqueNameAsync  name: {request.BalanceName}");
                var url = $"{_settings.BaseUrl}/{_settings.PostIsUniqueUrl}";
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var payload = new
                {
                    controllerName = _settings.ControllerName,
                    inputData = new
                    {
                        propertyName = "name",
                        id = 0,
                        value = request.BalanceName,
                        adjustmentId = "remotePlaceHolder",

                    }
                };

                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var response = await _httpClient.PostAsJsonAsync(url, payload, options);
                _logger.LogDebug($"Response from {url} {response}");
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                var result = JsonSerializer.Deserialize<bool>(json);

                _logger.LogInformation($"Unique check result for '{request.BalanceName}': {result}");

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to fetch verify for unique name from (name={request.BalanceName})");
                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"نام تراز یکتا نیست!" }
                        },
                    request.FolderPath
                    );
            }
        }

        public async Task<bool> PostFileAsync(string token, string file, APIRequestDto request)
        {
            try
            {
                _logger.LogInformation($"Starting PostFileAsync fileName: {request.FileName}");
                var url = $"{_settings.BaseUrl}/{_settings.PostBalanceSheetUrl}";
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var payload = new
                {
                    balancesheetType = 2,
                    caption = "تعریف فایل تراز جدید",
                    currencyType = 1,
                    file = file,
                    fileName = request.FileName,
                    fileType = 9,
                    hideBodyForm = false,
                    name = request.BalanceName,
                    showCancel = true
                };

                _logger.LogDebug($"payload: {payload}");

                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var response = await _httpClient.PostAsJsonAsync(url, payload, options);
                _logger.LogDebug($"Response from {url}: {response}");
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                _logger.LogDebug($"json: {json}");

                using var doc = JsonDocument.Parse(json);
                _logger.LogDebug($"doc: {doc.ToString}");

                if (doc.RootElement.TryGetProperty("models", out var models) && models.GetArrayLength() > 0)
                {
                    _logger.LogInformation("File uploaded successfully.");
                    return true;
                }
                else
                {
                    throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"پاسخ دریافتی از سرور معتبر نیست " }
                        },
                    request.FolderPath
                    );
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to post file (fileName: {request.FileName})");
                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $"خطا در ارتباط با سرور هنگام ارسال فایل" }
                    },
                request.FolderPath
                );
            }
        }  
    }
}
