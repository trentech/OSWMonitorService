using OSWMontiorService;
using Serilog;
using Serilog.Events;

Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().MinimumLevel.Override("Microsoft", LogEventLevel.Information).Enrich.FromLogContext()
    .WriteTo.File(Config.PATH + @"\log.txt", outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.Console(outputTemplate: "[{Timestamp:MM/dd/yyyy h:mm tt}] [{Level}] {Message:lj}{NewLine}{Exception}")
    .WriteTo.EventLog("OSW Monitoring Service")
    .CreateLogger();

IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<Worker>();
    }).UseSerilog().Build();


await host.RunAsync();
