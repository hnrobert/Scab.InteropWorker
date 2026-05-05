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
                listen.Protocols = HttpProtocols.Http2; // gRPC (MagicOnion)
            });
            options.ListenLocalhost(port + 1, listen =>
            {
                listen.Protocols = HttpProtocols.Http1; // REST API
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

        // HTTP API: export CAD file thumbnail via Inventor COM
        app.MapGet("/thumbnail", (string path, int? size, HttpContext ctx) =>
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return Results.BadRequest("File not found: " + path);

            var s = size ?? 256;
            var session = ctx.RequestServices.GetRequiredService<InventorSession>();

            var ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext is not ".ipt" and not ".iam" and not ".idw" and not ".ipn")
                return Results.BadRequest("Unsupported file type. Supported: .ipt, .iam, .idw, .ipn");

            var outputPath = Path.Combine(Path.GetTempPath(), "scab-worker-output", $"{Guid.NewGuid()}.png");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            try
            {
                session.EnsureConnected();
                session.ExportToPng(path, outputPath, s, s);
                var bytes = File.ReadAllBytes(outputPath);
                return Results.File(bytes, "image/png");
            }
            catch (Exception ex)
            {
                return Results.Problem(
                    $"Inventor export failed: {ex.Message}. Ensure Autodesk Inventor is installed.",
                    statusCode: 503);
            }
        });

        var lifetime = app.Services.GetRequiredService<IHostApplicationLifetime>();
        lifetime.ApplicationStopping.Register(() =>
        {
            Log.Information("Releasing Inventor COM session...");
            app.Services.GetRequiredService<InventorSession>().Dispose();
        });

        Log.Information("Scab.InteropWorker listening on 127.0.0.1:{Port} (gRPC) and :{HttpPort} (HTTP)", port, port + 1);
        Log.Information("Press Ctrl+C to stop");

        app.Run();

        Log.Information("Stopped");
    }
}
