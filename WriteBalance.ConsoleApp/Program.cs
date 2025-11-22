
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Newtonsoft.Json;
using WriteBalance.Application.Handlers;
using WriteBalance.Application.Interfaces;
using WriteBalance.Infrastructure.Config;
using WriteBalance.Infrastructure.Context;
using WriteBalance.Common.Logging;
using WriteBalance.Infrastructure.Repositories;
using WriteBalance.Infrastructure.Services;
using WriteBalanceConsoleApp;
class Program
{
    public static async Task Main(string[] args)
    {
        try
        {
            Logger.WriteEntry(JsonConvert.SerializeObject("Starting Applicaion"), $"Program:Main--typeReport:Info");

            var config = await InfoFileReader.ReadAsync(args);

            string folderName = config["of"];
            string path = config["op"];

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            if (!Directory.Exists($"{path}/{folderName}"))
                Directory.CreateDirectory($"{path}/{folderName}");


            Logger.WriteEntry(JsonConvert.SerializeObject($" AddressAPI :{config["AddressAPI"]} , UserNameAPI :{config["UserNameAPI"]}, AddressServer :{config["AddressServer"]}, DataBaseName :{config["DataBaseName"]}, UserName :{config["UserName"]}"), $"Program:Main --typeReport:Debug");

            using IHost host = Host.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((context, configBuilder) =>
                {
                    var basePath = AppContext.BaseDirectory;
                    configBuilder.SetBasePath(basePath);
                    configBuilder.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    configBuilder.AddEnvironmentVariables();
                    configBuilder.AddUserSecrets<Program>(optional: true);
                })
                .ConfigureServices((context, services) =>
                {
                    string connectionString = $"Server={config["AddressServer"]};Database={config["DataBaseName"]};User Id={config["UserName"]};Password={config["Password"]};TrustServerCertificate=True;";

                    Logger.WriteEntry(JsonConvert.SerializeObject($"connectionString: {connectionString}"), $"Program:Main --typeReport:Debug");
                    services.AddDbContext<AppDbContext>(options =>
                        options.UseSqlServer(connectionString));

                    //string bankConnectionString = $"Server={config["AddressServerBank"]};Database={config["DataBaseNameBank"]};User Id={config["UserNameBank"]};Password={config["PasswordBank"]};TrustServerCertificate=True;";
                    string bankConnectionString = $"Server={config["AddressServerBank"]};Database={config["DataBaseNameBank"]};Trusted_Connection=True;TrustServerCertificate=True;";

                    //string bankConnectionString = $"Server={config["AddressServer"]};Database={config["DataBaseName"]};User Id={config["UserName"]};Password={config["Password"]};TrustServerCertificate=True;";

                    services.AddDbContext<BankDbContext>(options =>
                        options.UseSqlServer(bankConnectionString));

                    services.AddDbContext<RayanBankDbContext>(options =>
                        options.UseSqlServer(bankConnectionString));

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
                    services.AddScoped<ICheckInput, CheckInput>();
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
                    services.AddSingleton<Logger>();

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
            Logger.WriteEntry(JsonConvert.SerializeObject($"Unhandled exception occurred in Main() : {ex}"), $"Program:Main --typeReport:Error");
            Environment.ExitCode = 604;

        }
    }
}
