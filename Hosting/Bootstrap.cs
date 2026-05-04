using Scab.InteropWorker.Interop;
using Serilog;
using Microsoft.AspNetCore.Server.Kestrel.Core;

namespace Scab.InteropWorker.Hosting;

public static class Bootstrap
{
    public static void Run(string[] args)
    {
        var port = 5100;
        if (args.Length > 0 && int.TryParse(args[0], out var p))
            port = p;

        Log.Logger = new LoggerConfiguration()
            .Enrich.FromLogContext()
            .MinimumLevel.Debug()
            .WriteTo.Console(
                theme: Serilog.Sinks.SystemConsole.Themes.AnsiConsoleTheme.Code,
                outputTemplate: "[{Timestamp:HH:mm:ss}] {Level} | {SourceContext} {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        var builder = WebApplication.CreateBuilder(args);

        builder.WebHost.ConfigureKestrel(options =>
        {
            options.ListenLocalhost(port, listen =>
            {
                listen.Protocols = HttpProtocols.Http2;
            });
        });

        builder.Services.AddLogging(logging =>
        {
            logging.ClearProviders();
            logging.AddSerilog();
        });

        builder.Services.AddMagicOnion();

        builder.Services.AddSingleton<InventorSession>();
        builder.Services.AddSingleton<Services.WorkerDispatchService>();

        var app = builder.Build();

        app.MapMagicOnionService();

        var lifetime = app.Services.GetRequiredService<IHostApplicationLifetime>();
        lifetime.ApplicationStopping.Register(() =>
        {
            Log.Information("Releasing Inventor COM session...");
            app.Services.GetRequiredService<InventorSession>().Dispose();
        });

        Log.Information("Scab.InteropWorker listening on 127.0.0.1:{Port} (HTTP/2)", port);
        Log.Information("Press Ctrl+C to stop");

        app.Run();

        Log.Information("Stopped");
    }
}
