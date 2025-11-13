using Azure.Core;
using Microsoft.Extensions.Logging;
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
using WriteBalance.Infrastructure.Config;

namespace WriteBalance.Infrastructure.Services
{
    public class AuthService : IAuthService
    {
        private readonly HttpClient _httpClient;
        private readonly AuthConfig _settings;

        private readonly ILogger<AuthService> _logger;

        public AuthService(HttpClient httpClient, AuthConfig settings, ILogger<AuthService> logger)
        {
            _httpClient = httpClient;
            _settings = settings;
            _logger = logger;
        }

        public async Task<string> GetAccessTokenAsync(APIRequestDto request, int companyId)
        {
            try
            {
                _logger.LogInformation("Starting GetAccessTokenAsync...");
                var url = $"{request.BaseUrl}/{_settings.AuthEndpointUrl}";

                var payload = new
                {
                    userName = request.UserNameAPI,
                    password = request.PasswordAPI,
                    companyId = companyId,
                    periodId = request.PeriodId
                };
                _logger.LogInformation($"payload:{payload}");

                var response = await _httpClient.PostAsJsonAsync(url, payload);
                Console.WriteLine(response);
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync();
                Console.WriteLine(json);
                var token = JsonSerializer.Deserialize<string>(json);

                return token!;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Failed to fetch Token from (username:{request.UserNameAPI}) , (companyId:{companyId})  and (periodId:{request.PeriodId})");
                throw new ConnectionMessageException(new ConnectionMessage
                {
                    MessageType = MessageType.Error,
                    Messages = new List<string> { "احراز هویت در اکسیر ناموفق  بوده است." }
                },
                request.FolderPath
                );
            }

        }
    }
}
