
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using WriteBalance.Application.Handlers;
using WriteBalance.Application.Interfaces;
using WriteBalance.Application.Interfaces.Repository;
using WriteBalance.Common.Logging;
using WriteBalance.Infrastructure.Config;
using WriteBalance.Infrastructure.Context;
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
            // جداسازی متغیرهای دریافتی کنسول
            var config = await InfoFileReader.ReadAsync(args);

            // تنظیم و چک کردن مسیر و نام پوشه وارد شده
            // در صورتی که وجود نداشته باشد ، ایجاد میشود
            string folderName = config["of"];
            string path = config["op"];

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            if (!Directory.Exists($"{path}/{folderName}"))
                Directory.CreateDirectory($"{path}/{folderName}");


            Logger.WriteEntry(JsonConvert.SerializeObject($" AddressAPI :{config["AddressAPI"]} , UserNameAPI :{config["UserNameAPI"]}, AddressServer :{config["AddressServer"]}, DataBaseName :{config["DataBaseName"]}, UserName :{config["UserName"]}"), $"Program:Main --typeReport:Debug");

            using IHost host = Host.CreateDefaultBuilder(args)
                .ConfigureLogging(logging =>
                {
                    logging.ClearProviders();
                })
                .ConfigureAppConfiguration((context, configBuilder) =>
                {
                    var basePath = AppContext.BaseDirectory;
                    configBuilder.SetBasePath(basePath);
                    // خواندن اطلاعات  از فایل 
                    //appsettings.json
                    configBuilder.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    configBuilder.AddEnvironmentVariables();
                    configBuilder.AddUserSecrets<Program>(optional: true);
                })
                .ConfigureServices((context, services) =>
                {                  
                   
                    string connectionString = $"Server={config["AddressServer"]};Database={config["DataBaseName"]};User Id={config["UserName"]};Password={config["Password"]};TrustServerCertificate=True;";
                    string connectionStringModules = $"Server={config["AddressServer"]};Database={config["ModulesDataBaseName"]};User Id={config["ModulesUserName"]};Password={config["ModulesPassword"]};TrustServerCertificate=True;";

                    Logger.WriteEntry(JsonConvert.SerializeObject($"connectionString: {connectionString},\n  connectionStringModules: {connectionStringModules}") , $"Program:Main --typeReport:Debug");

                    services.AddDbContext<AppDbContext>(options =>
                    {
                        options.UseSqlServer(connectionString);
                        options.EnableSensitiveDataLogging(false);
                        options.EnableDetailedErrors(false);
                    });


                    //Modules DB
                    services.AddDbContext<ModulesDbContext>(options =>
                    {
                        options.UseSqlServer(connectionStringModules);
                        options.EnableSensitiveDataLogging(false);
                        options.EnableDetailedErrors(false);
                    });

                    ////////////////////////////////////////////


                    //دیتابیس برای سما و همراه و کاربردی 
                    services.AddDbContext<BankDbContext>(options =>
                    {
                        options.UseSqlServer(connectionString);
                        options.EnableSensitiveDataLogging(false);
                        options.EnableDetailedErrors(false);
                    });

                   // دیتابیس برای رایان
                   services.AddDbContext<RayanBankDbContext>(options =>
                   {
                       options.UseSqlServer(connectionString);
                       options.EnableSensitiveDataLogging(false);
                       options.EnableDetailedErrors(false);
                   });

                   // دیتابیس برای پویا
                   services.AddDbContext<PouyaBankDbContext>(options =>
                   {
                       options.UseSqlServer(connectionString);
                       options.EnableSensitiveDataLogging(false);
                       options.EnableDetailedErrors(false);
                   });

                   // دیتابیس برای GL
                   services.AddDbContext<GLBankDbContext>(options =>
                   {
                       options.UseSqlServer(connectionString);
                       options.EnableSensitiveDataLogging(false);
                       options.EnableDetailedErrors(false);
                   });


                    ////////////////////////////////////////////

                    /*
                    string bankConnectionString = $"Server={config["AddressServerBank"]};Database={config["DataBaseNameBank"]};Trusted_Connection=True;TrustServerCertificate=True;";

                    services.AddDbContext<BankDbContext>(options =>
                      options.UseSqlServer(bankConnectionString));

                    services.AddDbContext<RayanBankDbContext>(options =>
                      options.UseSqlServer(bankConnectionString));

                    services.AddDbContext<PouyaBankDbContext>(options =>
                      options.UseSqlServer(bankConnectionString));

                    services.AddDbContext<GLBankDbContext>(options =>
                        options.UseSqlServer(bankConnectionString));

                    */
                    ////////////////////////////////////////////


                    services.Configure<AuthConfig>(
                        context.Configuration.GetSection("AuthConfig")
                    );
                    services.AddSingleton(resolver =>
                        resolver.GetRequiredService<Microsoft.Extensions.Options.IOptions<AuthConfig>>().Value
                    );
                    services.AddSingleton<AuthService>();

                    services.AddHttpClient<IAuthService, AuthService>();
                    services.AddHttpClient<IApiService, ApiService>();
                    services.AddScoped<IExcelExporter, ExcelExporter>();
                    services.AddScoped<ICheckInput, CheckInput>();
                    services.AddScoped<IBalanceGenerator, BalanceGenerator>();
                    services.AddScoped<IPouyaBalanceGenerator, PouyaBalanceGenerator>();
                    services.AddScoped<IGLBalanceGenerator, GLBalanceGenerator>();
                    services.AddScoped<IAllBalanceGenerator, AllBalanceGenerator>();             
                    services.AddScoped<IRayanBalanceGenerator, RayanBalanceGenerator>();
                    services.AddScoped<IFinancialRepository, FinancialRepository>();
                    services.AddScoped<IPooyaCodingRepository, PooyaCodingRepository>();
                    services.AddScoped<IPeriodRepository, PeriodRepository>();
                    services.AddScoped<ICompareBalance, CompareBalance>();
                    services.AddScoped<IFileEncoder, FileEncoder>();

                    // دریافت تنظیمات از فایل 
                    //appsettings.json
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
                    services.AddScoped<BalanceMerge>();
                    services.AddScoped<BalanceCheck>();
                    services.AddScoped<CalculateNewRows>();
                    services.AddScoped<BalanceController>();
                    services.AddScoped<CheckCodingPouya>();
                    services.AddScoped<GetPooyaCoding>();
                    services.AddScoped<Logger>();
                    

                })
                .Build();

            using (var scope = host.Services.CreateScope())
            {
                var controller = scope.ServiceProvider.GetRequiredService<BalanceController>();
                await controller.InputBalanceController(config);
            }
        }
        catch (Exception ex) // هر خطای غیر قابل پیش بینی رخ دهد  کد 604 برمیگردد
        {
            Logger.WriteEntry(JsonConvert.SerializeObject($"Unhandled exception occurred in Main() : {ex}"), $"Program:Main --typeReport:Error");
            Environment.ExitCode = 604;

        }
    }
}
