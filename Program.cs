using OSWMontiorService;

IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices(services =>
    {
        services.AddHostedService<DBTest>();
    })
    .Build();

await host.RunAsync();
