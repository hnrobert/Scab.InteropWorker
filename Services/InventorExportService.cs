using System.Collections.Immutable;
using System.Runtime.InteropServices;
using MagicOnion;
using MagicOnion.Server;
using Scab.InteropWorker.Interop;
using Scab.ServerD.Shared.Worker.Services;
using Scab.ServerD.Shared.Worker.Types;

namespace Scab.InteropWorker.Services;

public class InventorExportService(
    InventorSession inventor,
    ILogger<InventorExportService> logger)
    : ServiceBase<IWorkerDispatchService>, IWorkerDispatchService
{
    public async UnaryResult<WorkerJobResponse> SubmitJobAsync(WorkerJobRequest request)
    {
        if (request.JobType != EWorkerJobType.InventorPng)
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: "Unsupported job type");

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: "Inventor COM requires Windows");

        if (!File.Exists(request.FilePath))
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: $"File not found: {request.FilePath}");

        logger.LogInformation("Processing Inventor PNG export: {File}", request.FilePath);

        try
        {
            var (width, height) = ResolveDimensions(request.Parameters);
            var outputPath = Path.Combine(
                Path.GetTempPath(), "scab-worker-output",
                $"{Guid.NewGuid()}.png");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            inventor.EnsureConnected();
            inventor.ExportToPng(request.FilePath, outputPath, width, height);

            logger.LogInformation("Exported {File} → {Output} ({W}×{H})", request.FilePath, outputPath, width, height);
            return new WorkerJobResponse(request.JobId, 2, outputPath);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Inventor export failed for {File}", request.FilePath);
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: ex.Message);
        }
    }

    public async UnaryResult<WorkerJobResponse> GetJobStatusAsync(WorkerJobRequest request)
    {
        await Task.CompletedTask;
        return new WorkerJobResponse(request.JobId, 2);
    }

    private static (int Width, int Height) ResolveDimensions(ImmutableDictionary<string, string>? parameters)
    {
        var width = 256;
        var height = 256;

        if (parameters is not null)
        {
            if (parameters.TryGetValue("width", out var w) && int.TryParse(w, out var pw) && pw > 0)
                width = pw;
            if (parameters.TryGetValue("height", out var h) && int.TryParse(h, out var ph) && ph > 0)
                height = ph;
        }

        return (width, height);
    }
}
