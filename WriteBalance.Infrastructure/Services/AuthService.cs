using Azure.Core;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
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
    public class AuthService : IAuthService
    {
        private readonly HttpClient _httpClient;
        private readonly AuthConfig _settings;


        public AuthService(HttpClient httpClient, AuthConfig settings, ILogger<AuthService> logger)
        {
            _httpClient = httpClient;
            _settings = settings;
        }

        public async Task<string> GetAccessTokenAsync(APIRequestDto request, int companyId)
        {
            try
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Starting GetAccessTokenAsync"), $"AuthService:GetAccessTokenAsync --typeReport:Info");

                var url = $"{request.BaseUrl}/{_settings.AuthEndpointUrl}";

                var payload = new
                {
                    userName = request.UserNameAPI,
                    password = request.PasswordAPI,
                    companyId = companyId,
                    periodId = request.PeriodId
                };

                Logger.WriteEntry(JsonConvert.SerializeObject($"payload:{payload}"), $"AuthService:GetAccessTokenAsync --typeReport:Info");

                var response = await _httpClient.PostAsJsonAsync(url, payload);
                response.EnsureSuccessStatusCode();

                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Login failed. Status: {response.StatusCode}, Error: {error}");
                }

                var json = await response.Content.ReadAsStringAsync();
                Console.WriteLine(json);
                var token = System.Text.Json.JsonSerializer.Deserialize<string>(json);

                return token!;
            }
            catch (Exception ex)
            {
                Logger.WriteEntry(JsonConvert.SerializeObject($"Failed to fetch Token from (username:{request.UserNameAPI}) , (companyId:{companyId})  and (periodId:{request.PeriodId})"), $"AuthService:GetAccessTokenAsync --typeReport:Error");
                Logger.WriteEntry(JsonConvert.SerializeObject(ex), $"AuthService:GetAccessTokenAsync --typeReport:Error");
                throw new ConnectionMessageException(new ConnectionMessage
                {
                    MessageType = MessageType.Error,
                    Messages = new List<string> { "احراز هویت اکسیر ناموفق  ." }
                },
                request.FolderPath
                );
            }

        }
    }
}
