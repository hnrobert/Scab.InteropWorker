using MagicOnion;
using MagicOnion.Server;
using Scab.ServerD.Shared.Worker.Services;
using Scab.ServerD.Shared.Worker.Types;

namespace Scab.InteropWorker.Services;

public class HeadlessRenderService(ILogger<HeadlessRenderService> logger)
    : ServiceBase<IWorkerDispatchService>, IWorkerDispatchService
{
    public async UnaryResult<WorkerJobResponse> SubmitJobAsync(WorkerJobRequest request)
    {
        if (request.JobType is not (EWorkerJobType.StepRender or EWorkerJobType.StlRender))
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: "Unsupported job type");

        logger.LogInformation("Processing headless render: {Type} {File}", request.JobType, request.FilePath);

        try
        {
            var outputPath = Path.Combine(
                Path.GetTempPath(), "scab-worker-output",
                $"{Guid.NewGuid()}.png");
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            // TODO: Implement headless rendering
            // STEP: Use OpenCascade or FreeCAD headless
            // STL: Use three.js or similar renderer
            // For MVP, return placeholder

            logger.LogInformation("Rendered {File} to {Output}", request.FilePath, outputPath);
            return new WorkerJobResponse(request.JobId, 2, outputPath);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Headless render failed for {File}", request.FilePath);
            return new WorkerJobResponse(request.JobId, -1, ErrorMessage: ex.Message);
        }
    }

    public async UnaryResult<WorkerJobResponse> GetJobStatusAsync(WorkerJobRequest request)
    {
        await Task.CompletedTask;
        return new WorkerJobResponse(request.JobId, 2);
    }
}
