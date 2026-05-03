using System.Collections.Immutable;
using System.Runtime.InteropServices;
using MagicOnion;
using MagicOnion.Server;
using Scab.ServerD.Shared.Worker.Services;
using Scab.ServerD.Shared.Worker.Types;

namespace Scab.InteropWorker.Services;

public class InventorExportService(ILogger<InventorExportService> logger)
    : ServiceBase<IWorkerDispatchService>, IWorkerDispatchService
{
    public async UnaryResult<WorkerJobResponse> SubmitJobAsync(WorkerJobRequest request)
    {
        if (request.JobType != EWorkerJobType.InventorPng)
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: "Unsupported job type");

        logger.LogInformation("Processing Inventor export: {File}", request.FilePath);

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: "Inventor COM requires Windows");

        try
        {
            var outputPath = await ExportToPng(request.FilePath, request.Parameters);
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

    private async Task<string> ExportToPng(string filePath, ImmutableDictionary<string, string>? parameters)
    {
        var outputPath = Path.Combine(
            Path.GetTempPath(), "scab-worker-output",
            $"{Guid.NewGuid()}.png");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // COM interop with Inventor
        // Requires Autodesk Inventor installed on the machine
        var width = 256;
        var height = 256;

        logger.LogInformation("Exporting {File} to {Output} ({W}x{H})", filePath, outputPath, width, height);

        // TODO: Implement COM automation via InventorComInterop
        // var app = InventorComInterop.GetOrCreateInventorInstance();
        // var doc = app.Documents.Open(filePath, false);
        // doc.SaveAsBitmap(outputPath, width, height);
        // doc.Close(true);

        await Task.CompletedTask;
        return outputPath;
    }
}
