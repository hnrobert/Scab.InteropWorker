using System.Collections.Immutable;
using System.Runtime.InteropServices;
using MagicOnion;
using MagicOnion.Server;
using Scab.InteropWorker.Interop;
using Scab.ServerD.Shared.Worker.Services;
using Scab.ServerD.Shared.Worker.Types;

namespace Scab.InteropWorker.Services;

public class WorkerDispatchService(
    InventorSession inventor,
    ILogger<WorkerDispatchService> logger)
    : ServiceBase<IWorkerDispatchService>, IWorkerDispatchService
{
    public async UnaryResult<WorkerJobResponse> SubmitJobAsync(WorkerJobRequest request)
    {
        logger.LogInformation("Received job {JobId}: {Type} {File}", request.JobId, request.JobType, request.FilePath);

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: "COM interop requires Windows");

        if (!File.Exists(request.FilePath))
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: $"File not found: {request.FilePath}");

        try
        {
            return request.JobType switch
            {
                EWorkerJobType.InventorPng => await HandleInventorPng(request),
                EWorkerJobType.StepRender or EWorkerJobType.StlRender => await HandleHeadlessRender(request),
                _ => new WorkerJobResponse(request.JobId, -1, ErrorMessage: $"Unknown job type: {request.JobType}")
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Job {JobId} failed", request.JobId);
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: ex.Message);
        }
    }

    public async UnaryResult<WorkerJobResponse> GetJobStatusAsync(WorkerJobRequest request)
    {
        await Task.CompletedTask;
        return new WorkerJobResponse(request.JobId, 2);
    }

    private async UnaryResult<WorkerJobResponse> HandleInventorPng(WorkerJobRequest request)
    {
        var (width, height) = ResolveDimensions(request.Parameters);
        var outputPath = Path.Combine(
            Path.GetTempPath(), "scab-worker-output",
            $"{Guid.NewGuid()}.png");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        inventor.EnsureConnected();
        inventor.ExportToPng(request.FilePath, outputPath, width, height);

        logger.LogInformation("Inventor export: {File} → {Output} ({W}×{H})",
            request.FilePath, outputPath, width, height);

        return new WorkerJobResponse(request.JobId, 2, outputPath);
    }

    private async UnaryResult<WorkerJobResponse> HandleHeadlessRender(WorkerJobRequest request)
    {
        var outputPath = Path.Combine(
            Path.GetTempPath(), "scab-worker-output",
            $"{Guid.NewGuid()}.png");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // TODO: Implement headless rendering (OpenCascade/FreeCAD for STEP, three.js for STL)

        logger.LogInformation("Headless render ({Type}): {File} → {Output}", request.JobType, request.FilePath, outputPath);

        return new WorkerJobResponse(request.JobId, 2, outputPath);
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
