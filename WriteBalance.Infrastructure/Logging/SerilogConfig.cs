using Microsoft.Extensions.Hosting;
using Serilog;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Infrastructure.Logging
{
    public static class SerilogConfig
    {
        public static void ConfigureSerilog(HostBuilderContext context, LoggerConfiguration loggerConfiguration)
        {
            /*
            var env = context.HostingEnvironment;
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var exeDirectory = AppContext.BaseDirectory;

            var logDirectory = Path.Combine(exeDirectory, "logs");
            Directory.CreateDirectory(logDirectory);

            var logFilePath = Path.Combine(logDirectory, $"log-{timestamp}.txt");

            loggerConfiguration
                .ReadFrom.Configuration(context.Configuration)
                .Enrich.FromLogContext()
                .Enrich.WithProperty("Environment", env.EnvironmentName)
                .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
                .WriteTo.Console()
                .WriteTo.File(
                    path: logFilePath,
                    restrictedToMinimumLevel: LogEventLevel.Information,
                    shared: true);
            */
            var env = context.HostingEnvironment;
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var exeDirectory = AppContext.BaseDirectory;
            var logDirectory = Path.Combine(exeDirectory, "logs");
            Directory.CreateDirectory(logDirectory);

            var logFilePath = Path.Combine(logDirectory, $"log-{timestamp}.txt");

            loggerConfiguration
                .ReadFrom.Configuration(context.Configuration)
                .Enrich.FromLogContext()
                .Enrich.WithProperty("Environment", env.EnvironmentName)
                .MinimumLevel.Verbose()
                .MinimumLevel.Override("Microsoft", LogEventLevel.Verbose)
                .MinimumLevel.Override("System", LogEventLevel.Verbose)
                .WriteTo.Console(restrictedToMinimumLevel: LogEventLevel.Verbose)
                .WriteTo.File(
                    path: logFilePath,
                    restrictedToMinimumLevel: LogEventLevel.Verbose,
                    rollingInterval: RollingInterval.Infinite,
                    shared: true);

        }
    }
}
