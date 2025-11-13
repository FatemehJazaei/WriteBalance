
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using WriteBalance.Application.DTOs;
using WriteBalance.Application.Handlers;
using WriteBalance.Application.Interfaces;
using WriteBalance.Infrastructure.Config;
using WriteBalance.Infrastructure.Context;
using WriteBalance.Infrastructure.Logging;
using WriteBalance.Infrastructure.Repositories;
using WriteBalance.Infrastructure.Services;
using WriteBalance.ConsoleApp;
class Program
{
    public static async Task Main(string[] args)
    {

        Log.Logger = new LoggerConfiguration()
            .WriteTo.File("logs/log.txt", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        try
        {

            Log.Information("Starting Applicaion ... ");

            var config = await InfoFileReader.ReadAsync(args);
            Log.Information($"config:{config.Keys}, {config.Values}");

            using IHost host = Host.CreateDefaultBuilder(args)
                .UseSerilog((context, loggerConfig) =>
                {
                   SerilogConfig.ConfigureSerilog(context, loggerConfig);
                })
                .ConfigureAppConfiguration((context, configBuilder) =>
                {
                    configBuilder.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    configBuilder.AddUserSecrets<Program>(optional: true);
                })
                .ConfigureServices((context, services) =>
                {
                    string connectionString = $"Server={config["AddressServer"]};Database={config["DataBaseName"]};User Id={config["UserName"]};Password={config["Password"]};TrustServerCertificate=True;";
                    services.AddDbContext<AppDbContext>(options =>
                        options.UseSqlServer(connectionString));

                    //string bankConnectionString = $"Server={config["AddressServerBank"]};Database={config["DataBaseNameBank"]};User Id={config["UserNameBank"]};Password={config["PasswordBank"]};TrustServerCertificate=True;";
                    string bankConnectionString = $"Server={config["AddressServerBank"]};Database={config["DataBaseNameBank"]};Trusted_Connection=True;TrustServerCertificate=True;";
                    services.AddDbContext<BankDbContext>(options =>
                        options.UseSqlServer(bankConnectionString));

                    //string bankConnectionString = $"Server={config["AddressServerBank"]};Database={config["DataBaseNameBank"]};User Id={config["UserNameBank"]};Password={config["PasswordBank"]};TrustServerCertificate=True;";
                    string rayanConnectionString = $"Server={config["AddressServerBank"]};Database={config["DataBaseNameBank"]};Trusted_Connection=True;TrustServerCertificate=True;";
                    services.AddDbContext<RayanBankDbContext>(options =>
                        options.UseSqlServer(rayanConnectionString));

                    services.Configure<AuthConfig>(
                        context.Configuration.GetSection("AuthConfig")
                    );
                    services.AddSingleton(resolver =>
                        resolver.GetRequiredService<Microsoft.Extensions.Options.IOptions<AuthConfig>>().Value
                    );
                    services.AddSingleton<AuthService>();

                    services.AddHttpClient<IAuthService, AuthService>();
                    services.AddHttpClient<IApiService, ApiService>();
                    services.AddSingleton<IExcelExporter, ExcelExporter>();
                    services.AddSingleton<IBalanceGenerator, BalanceGenerator>();
                    services.AddScoped<IFinancialRepository, FinancialRepository>();
                    services.AddScoped<IPeriodRepository, PeriodRepository>();
                    services.AddScoped<IFileEncoder, FileEncoder>();

                    var apiSettings = context.Configuration.GetSection("ApiSettings").Get<ApiSettings>()!;
                    var apiConfig = new ApiConfig
                    {
                        BaseUrl = config["AddressAPI"],
                        PostIsUniqueUrl = apiSettings.PostIsUniqueUrl,
                        PostBalanceSheetUrl = apiSettings.PostBalanceSheetUrl,
                        ControllerName = apiSettings.ControllerName,
                        RetryCount = apiSettings.RetryCount,
                        RetryDelaySeconds = apiSettings.RetryDelaySeconds,
                    };
                    services.AddSingleton(apiConfig);

                    var authSettings = context.Configuration.GetSection("AuthConfig").Get<AuthConfig>()!;
                    var authConfig = new AuthConfig
                    {
                        AuthEndpointUrl = authSettings.AuthEndpointUrl,
                    };
                    services.AddSingleton(authConfig);

                    services.AddScoped<WriteBalanceHandler>();
                    services.AddScoped<BalanceController>();

                })
                .Build();

            using (var scope = host.Services.CreateScope())
            {
                var controller = scope.ServiceProvider.GetRequiredService<BalanceController>();
                await controller.InputBalanceController(config);
            }
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, " Unhandled exception occurred in Main()");
            Environment.ExitCode = 604;
        }
        finally
        {
            Log.CloseAndFlush();
        }


    }
}
