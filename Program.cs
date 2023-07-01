using OSWMonitorService;
using OSWMonitorService.JSON;
using Serilog;
using Serilog.Events;
using Serilog.Sinks.SystemConsole.Themes;

try
{
    Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().MinimumLevel.Override("Microsoft", LogEventLevel.Information).Enrich.FromLogContext()
    .WriteTo.File(Config.PATH + @"\log.txt", outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.Console(outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm:ss tt}] [{Level:u3}] {Message:lj}{NewLine}{Exception}", theme : AnsiConsoleTheme.Code)
    .WriteTo.EventLog("OSW Monitoring Service", manageEventSource: true)
    .CreateLogger();
} catch
{
    Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().MinimumLevel.Override("Microsoft", LogEventLevel.Information).Enrich.FromLogContext()
    .WriteTo.File(Config.PATH + @"\log.txt", outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.Console(outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm:ss tt}] [{Level:u3}] {Message:lj}{NewLine}{Exception}", theme: AnsiConsoleTheme.Code)
    .WriteTo.EventLog(source: "Application")
    .CreateLogger();
}

IHost host = Host.CreateDefaultBuilder(args)
    .UseWindowsService()
    .ConfigureServices(services =>
    {
        services.AddHostedService<Worker>();
    }).UseSerilog().Build();


await host.RunAsync();
