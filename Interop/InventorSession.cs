using System.Runtime.InteropServices;

namespace Scab.InteropWorker.Interop;

public sealed class InventorSession : IDisposable
{
    private object? _app;
    private bool _createdInstance;
    private bool _disposed;
    private readonly Lock _lock = new();

    public bool IsConnected
    {
        get { lock (_lock) return _app != null; }
    }

    public void EnsureConnected()
    {
        lock (_lock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_app != null) return;

            var existing = InventorComInterop.TryGetActiveInstance();
            if (existing != null)
            {
                _app = existing;
                _createdInstance = false;
            }
            else
            {
                _app = InventorComInterop.CreateInstance();
                _createdInstance = true;
            }

            dynamic app = _app;
            app.SilentOperation = true;
        }
    }

    public string ExportToPng(string filePath, string outputPath, int width, int height)
    {
        lock (_lock)
        {
            EnsureConnected();

            dynamic app = _app!;
            dynamic? doc = null;

            try
            {
                doc = app.Documents.Open(filePath, false);
                if (doc is null)
                    throw new InvalidOperationException($"Inventor failed to open document: {filePath}");

                // kPartDocumentObject = 12290, kAssemblyDocumentObject = 12291
                var docType = (int)doc.DocumentType;
                if (docType is 12290 or 12291)
                {
                    dynamic camera = app.ActiveView.Camera;
                    camera.ViewOrientationType = 10763; // kIsoTopRightViewOrientation
                    camera.Fit();
                    camera.Apply();
                    Marshal.ReleaseComObject(camera);
                }

                app.ActiveView.SaveAsBitmap(outputPath, width, height);
                return outputPath;
            }
            finally
            {
                if (doc is not null)
                {
                    doc.Close(true);
                    Marshal.ReleaseComObject(doc);
                }
            }
        }
    }

    public void Dispose()
    {
        lock (_lock)
        {
            if (_disposed) return;
            _disposed = true;

            if (_app is null) return;

            try
            {
                if (_createdInstance)
                {
                    ((dynamic)_app).Quit();
                }
            }
            catch
            {
                // Best-effort cleanup
            }
            finally
            {
                Marshal.ReleaseComObject(_app);
                _app = null;
            }
        }
    }
}
