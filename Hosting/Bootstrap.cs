using Serilog;

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

        builder.WebHost.UseUrls($"http://127.0.0.1:{port}");

        builder.Services.AddLogging(logging =>
        {
            logging.ClearProviders();
            logging.AddSerilog();
        });

        builder.Services.AddMagicOnion();

        builder.Services.AddSingleton<Services.InventorExportService>();
        builder.Services.AddSingleton<Services.HeadlessRenderService>();

        var app = builder.Build();

        app.MapMagicOnionService();

        Log.Information("Scab.InteropWorker listening on 127.0.0.1:{Port}", port);
        Log.Information("Press Ctrl+C to stop");

        app.Run();

        Log.Information("Stopped");
    }
}
