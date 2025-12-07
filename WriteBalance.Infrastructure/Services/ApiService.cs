using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Newtonsoft.Json;
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
using WriteBalance.Common.Logging;
using WriteBalance.Infrastructure.Config;

namespace WriteBalance.Infrastructure.Services
{
    public class ApiService : IApiService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiConfig _settings;

        public ApiService(HttpClient httpClient, ApiConfig settings)
        {
            _httpClient = httpClient;
            _settings = settings;
        }

        public async Task<bool> GetVerifyUniqueNameAsync(string token, APIRequestDto request)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Starting GetVerifyUniqueNameAsync"), $"ApiService: GetVerifyUniqueNameAsync--typeReport:Info");

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

                response.EnsureSuccessStatusCode();
                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Login failed. Status: {response.StatusCode}, Error: {error}");
                }

                var json = await response.Content.ReadAsStringAsync();
                var result = System.Text.Json.JsonSerializer.Deserialize<bool>(json);

                if (result == false )
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Failed to fetch verify for unique name from (name={request.BalanceName})"), $"ApiService:GetVerifyUniqueNameAsync --typeReport:Error");
                    throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"نام تراز تکراری است!" }
                            },
                        request.FolderPath
                        );

                }
                return result;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Failed to fetch verify for unique name from (name={request.BalanceName})"), $"ApiService:GetVerifyUniqueNameAsync --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"ApiService:GetVerifyUniqueNameAsync --typeReport:Error");

                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"نام تراز تکراری است!" }
                        },
                    request.FolderPath
                    );
            }
        }

        public async Task<bool> PostFileAsync(string token, string file, APIRequestDto request)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Starting PostFileAsync fileName: {request.FileNameRial}"), $"ApiService: PostFileAsync--typeReport:Info");

                var url = $"{_settings.BaseUrl}/{_settings.PostBalanceSheetUrl}";
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var payload = new
                {
                    balancesheetType = 2,
                    caption = "تعریف فایل تراز جدید",
                    currencyType = 1,
                    file = file,
                    fileName = request.FileNameRial,
                    fileType = 9,
                    hideBodyForm = false,
                    name = request.BalanceName,
                    showCancel = true
                };

                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var response = await _httpClient.PostAsJsonAsync(url, payload, options);

                response.EnsureSuccessStatusCode();

                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    throw new Exception($"PostFileAsync failed. Status: {response.StatusCode}, Error: {error}");
                }

                var json = await response.Content.ReadAsStringAsync();

                using var doc = JsonDocument.Parse(json);
   
                if (doc.RootElement.TryGetProperty("models", out var models) && models.GetArrayLength() > 0)
                {

                    Logger.WriteEntry(JsonConvert.SerializeObject("File uploaded successfully."), $"ApiService: PostFileAsync--typeReport:Info");
                    return true;
                }
                else
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("Uplouding file failed!."), $"ApiService: PostFileAsync --typeReport:Error");
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
            catch(Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject("Uplouding file failed!."), $"ApiService:PostFileAsync --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"ApiService:PostFileAsync --typeReport:Error");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $"خطا در ارتباط با سرور - ارسال فایل ناموفق" }
                    },
                request.FolderPath
                );
            }
        }

        //Arzi
        public async Task<bool> GetVerifyUniqueNameArziAsync(string token, APIRequestDto request)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Starting GetVerifyUniqueNameArziAsync"), $"ApiService: GetVerifyUniqueNameArziAsync--typeReport:Info");

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

                response.EnsureSuccessStatusCode();
                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Login failed. Status: {response.StatusCode}, Error: {error}");
                }

                var json = await response.Content.ReadAsStringAsync();
                var result = System.Text.Json.JsonSerializer.Deserialize<bool>(json);

                if (result == false)
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject($"Failed to fetch verify for unique name from (name={request.BalanceName})"), $"ApiService:GetVerifyUniqueNameArziAsync --typeReport:Error");
                    throw new ConnectionMessageException(
                            new ConnectionMessage
                            {
                                MessageType = MessageType.Error,
                                Messages = new List<string> { $"نام تراز تکراری است!" }
                            },
                        request.FolderPath
                        );

                }
                return result;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Failed to fetch verify for unique name from (name={request.BalanceName})"), $"ApiService:GetVerifyUniqueNameArziAsync --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"ApiService:GetVerifyUniqueNameArziAsync --typeReport:Error");

                throw new ConnectionMessageException(
                        new ConnectionMessage
                        {
                            MessageType = MessageType.Error,
                            Messages = new List<string> { $"نام تراز تکراری است!" }
                        },
                    request.FolderPath
                    );
            }
        }

        public async Task<bool> PostFileArziAsync(string token, string file, APIRequestDto request)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Starting PostFileArziAsync fileName: {request.FileNameArzi}"), $"ApiService: PostFileArziAsync--typeReport:Info");

                var url = $"{_settings.BaseUrl}/{_settings.PostBalanceSheetUrl}";
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var payload = new
                {
                    balancesheetType = 2,
                    caption = "تعریف فایل تراز جدید",
                    currencyType = 2,
                    file = file,
                    fileName = request.FileNameArzi,
                    fileType = 9,
                    hideBodyForm = false,
                    name = request.BalanceName,
                    showCancel = true
                };

                var options = new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var response = await _httpClient.PostAsJsonAsync(url, payload, options);

                response.EnsureSuccessStatusCode();

                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    throw new Exception($"PostFileAsync failed. Status: {response.StatusCode}, Error: {error}");
                }

                var json = await response.Content.ReadAsStringAsync();

                using var doc = JsonDocument.Parse(json);

                if (doc.RootElement.TryGetProperty("models", out var models) && models.GetArrayLength() > 0)
                {

                    Logger.WriteEntry(JsonConvert.SerializeObject("File uploaded successfully."), $"ApiService: PostFileArziAsync--typeReport:Info");
                    return true;
                }
                else
                {
                    Logger.WriteEntry(JsonConvert.SerializeObject("Uplouding file failed!."), $"ApiService: PostFileArziAsync --typeReport:Error");
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
                Logger.WriteEntry(JsonConvert.SerializeObject("Uplouding file failed!."), $"ApiService:PostFileAsync --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"ApiService:PostFileAsync --typeReport:Error");

                throw new ConnectionMessageException(
                    new ConnectionMessage
                    {
                        MessageType = MessageType.Error,
                        Messages = new List<string> { $"خطا در ارتباط با سرور - ارسال فایل ناموفق" }
                    },
                request.FolderPath
                );
            }
        }
    }
}
