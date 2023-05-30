using OSWMonitorService;
using Serilog;
using Serilog.Events;

try
{
    Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().MinimumLevel.Override("Microsoft", LogEventLevel.Information).Enrich.FromLogContext()
    .WriteTo.File(Config.PATH + @"\log.txt", outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.Console(outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.EventLog("OSW Monitoring Service", manageEventSource: true)
    .CreateLogger();
} catch
{
    Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().MinimumLevel.Override("Microsoft", LogEventLevel.Information).Enrich.FromLogContext()
    .WriteTo.File(Config.PATH + @"\log.txt", outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.Console(outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.EventLog(source: "Application")
    .CreateLogger();
}

IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<Worker>();
    }).UseSerilog().Build();


await host.RunAsync();
