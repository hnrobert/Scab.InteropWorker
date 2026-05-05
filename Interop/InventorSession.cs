using System.Runtime.InteropServices;
using Serilog;

namespace Scab.InteropWorker.Interop;

/// <summary>
/// Manages a persistent Inventor COM instance on a dedicated STA thread
/// with a Windows message pump for proper COM callback processing.
/// Inventor is opened in user-visible mode.
/// </summary>
public sealed class InventorSession : IDisposable
{
    private object? _app;
    private bool _createdInstance;
    private bool _disposed;

    // Persistent STA thread with message pump
    private Thread? _staThread;
    private IntPtr _hwnd;
    private readonly List<WorkItem> _queue = [];
    private readonly Lock _queueLock = new();
    private uint _wmWork;
    private readonly ManualResetEventSlim _ready = new(false);
    private static readonly object s_classLock = new();
    private static string? s_registeredClassName;

    private record struct WorkItem(Delegate Action, TaskCompletionSource<object?> Tcs);

    #region Win32

    [DllImport("ole32.dll")]
    private static extern int OleInitialize(IntPtr pvReserved);

    [DllImport("ole32.dll")]
    private static extern void OleUninitialize();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern short RegisterClassEx(ref WNDCLASSEX lpwcx);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateWindowEx(uint ex, string cls, string name, uint style,
        int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr inst, IntPtr param);

    [DllImport("user32.dll")]
    private static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wp, IntPtr lp);

    [DllImport("user32.dll")]
    private static extern bool GetMessage(out MSG msg, IntPtr hWnd, uint min, uint max);

    [DllImport("user32.dll")]
    private static extern bool TranslateMessage(ref MSG msg);

    [DllImport("user32.dll")]
    private static extern IntPtr DispatchMessage(ref MSG msg);

    [DllImport("user32.dll")]
    private static extern uint RegisterWindowMessage(string name);

    [DllImport("user32.dll")]
    private static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wp, IntPtr lp);

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam, lParam;
        public uint time;
        public int pt_x, pt_y;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WNDCLASSEX
    {
        public uint cbSize;
        public uint style;
        public WndProcDelegate lpfnWndProc;
        public int cbClsExtra, cbWndExtra;
        public IntPtr hInstance, hIcon, hCursor, hbrBackground;
        public string? lpszMenuName, lpszClassName;
        public IntPtr hIconSm;
    }

    private delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wp, IntPtr lp);

    private static readonly WndProcDelegate s_wndProc = StaticWndProc;

    private static IntPtr StaticWndProc(IntPtr hWnd, uint msg, IntPtr wp, IntPtr lp)
    {
        // The work message is dispatched per-instance via the queue
        // We use a static dispatch since Win32 WNDPROC is static
        if (msg >= 0xC000) // registered message
        {
            lock (s_instances)
            {
                foreach (var inst in s_instances)
                {
                    if (inst._hwnd == hWnd && msg == inst._wmWork)
                    {
                        inst.ProcessQueue();
                        return IntPtr.Zero;
                    }
                }
            }
        }
        return DefWindowProc(hWnd, msg, wp, lp);
    }

    private static readonly List<InventorSession> s_instances = [];

    #endregion

    private void StartStaThread()
    {
        _wmWork = RegisterWindowMessage("Scab.STA." + Environment.TickCount64);
        Log.Debug("Registered WM_WORK={WmWork}", _wmWork);

        _staThread = new Thread(() =>
        {
            Log.Debug("STA thread starting, apartment={Apt}", Thread.CurrentThread.GetApartmentState());
            var hr = OleInitialize(IntPtr.Zero);
            Log.Debug("OleInitialize hr={HR}", hr);

            var hInst = Marshal.GetHINSTANCE(GetType().Assembly.Modules.First());
            string className;
            lock (s_classLock)
            {
                s_registeredClassName ??= "ScabStaWnd_" + Environment.TickCount64;
                className = s_registeredClassName;
            }

            var wc = new WNDCLASSEX
            {
                cbSize = (uint)Marshal.SizeOf<WNDCLASSEX>(),
                lpfnWndProc = s_wndProc,
                hInstance = hInst,
                lpszClassName = className
            };
            var atom = RegisterClassEx(ref wc);
            Log.Debug("RegisterClassEx atom={Atom}, className={ClassName}, hInst={HInst}", atom, className, hInst);

            _hwnd = CreateWindowEx(0, className, "", 0, 0, 0, 0, 0,
                new IntPtr(-3), IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Log.Debug("CreateWindowEx hwnd={Hwnd}", _hwnd);

            lock (s_instances) s_instances.Add(this);

            Log.Information("STA thread ready, hwnd={Hwnd}, wmWork={WmWork}", _hwnd, _wmWork);
            _ready.Set();

            while (GetMessage(out var msg, IntPtr.Zero, 0, 0))
            {
                TranslateMessage(ref msg);
                DispatchMessage(ref msg);
            }

            lock (s_instances) s_instances.Remove(this);
            if (_hwnd != IntPtr.Zero) DestroyWindow(_hwnd);
            OleUninitialize();
        })
        {
            Name = "Inventor-STA",
            IsBackground = true,
        };
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();
    }

    private void ProcessQueue()
    {
        List<WorkItem> items;
        lock (_queueLock)
        {
            items = [.. _queue];
            _queue.Clear();
        }
        Log.Debug("ProcessQueue: {Count} items on thread {ThreadId}", items.Count, Environment.CurrentManagedThreadId);
        foreach (var item in items)
        {
            try
            {
                var result = item.Action.DynamicInvoke();
                item.Tcs.SetResult(result);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "ProcessQueue item failed");
                item.Tcs.SetException(ex.InnerException ?? ex);
            }
        }
    }

    public T RunOnSta<T>(Func<T> action, int timeoutMs = 30000)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (_staThread is null) StartStaThread();

        if (!_ready.Wait(5000))
            throw new TimeoutException("STA thread failed to start within 5s");

        var tcs = new TaskCompletionSource<object?>();
        lock (_queueLock)
        {
            _queue.Add(new WorkItem(new Func<object?>(() => action()), tcs));
        }

        Log.Debug("PostMessage hwnd={Hwnd} msg={Msg}", _hwnd, _wmWork);
        PostMessage(_hwnd, _wmWork, IntPtr.Zero, IntPtr.Zero);

        if (!tcs.Task.Wait(timeoutMs))
            throw new TimeoutException($"COM call timed out after {timeoutMs}ms");

        return (T)tcs.Task.Result!;
    }

    public bool IsConnected => Volatile.Read(ref _app) != null;

    public void EnsureConnected()
    {
        if (IsConnected) return;

        RunOnSta(() =>
        {
            if (_app != null) return true;

            Log.Information("Looking for running Inventor instance...");
            var existing = InventorComInterop.TryGetActiveInstance();
            if (existing != null)
            {
                _createdInstance = false;
                _app = existing;
                Log.Information("Attached to existing Inventor instance");
                return true;
            }

            Log.Information("No running Inventor found, creating new instance...");
            _createdInstance = true;
            _app = InventorComInterop.CreateInstance();
            dynamic app = _app;
            app.Visible = true;
            app.SilentOperation = true;
            Log.Information("Created new Inventor COM instance");
            return true;
        });
    }

    public string ExportToPng(string filePath, string outputPath, int width, int height)
    {
        EnsureConnected();
        return RunOnSta(() =>
        {
            if (_app == null) throw new InvalidOperationException("Inventor not connected.");

            Log.Information("ExportToPng: opening {Path}", filePath);
            dynamic app = _app;
            dynamic? doc = null;

            try
            {
                doc = app.Documents.Open(filePath, true);
                if (doc is null)
                    throw new InvalidOperationException($"Failed to open: {filePath}");

                var docType = (int)doc.DocumentType;
                Log.Debug("Document type: {DocType}", docType);

                if (docType is 12290 or 12291) // Part or Assembly
                {
                    var view = app.ActiveView;
                    if (view is null)
                        throw new InvalidOperationException("ActiveView is null after opening document.");

                    dynamic camera = view.Camera;
                    camera.Fit();
                    camera.Apply();
                    Marshal.ReleaseComObject(camera);

                    view.SaveAsBitmap(outputPath, width, height);
                }
                else
                {
                    app.ActiveView?.SaveAsBitmap(outputPath, width, height);
                }
                Log.Information("ExportToPng: saved to {Output}", outputPath);
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
        });
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_app is not null)
        {
            try
            {
                RunOnSta(() =>
                {
                    try { if (_createdInstance && _app is not null) ((dynamic)_app).Quit(); }
                    catch { }
                    if (_app is not null) Marshal.ReleaseComObject(_app);
                    _app = null;
                    return true;
                }, timeoutMs: 5000);
            }
            catch { }
        }
    }
}
