using System.Runtime.InteropServices;

namespace Scab.InteropWorker.Interop;

public static class InventorComInterop
{
    internal const string ProgId = "Inventor.Application";

    [DllImport("oleaut32.dll", PreserveSig = false)]
    [return: MarshalAs(UnmanagedType.IUnknown)]
    private static extern object GetActiveObject(
        [In, MarshalAs(UnmanagedType.LPStruct)] Guid clsid,
        IntPtr reserved);

    public static void EnsureWindows()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            throw new PlatformNotSupportedException("Inventor COM interop requires Windows.");
    }

    public static Type GetInventorType()
    {
        return Type.GetTypeFromProgID(ProgId)
               ?? throw new InvalidOperationException("Autodesk Inventor is not installed or not registered for COM interop.");
    }

    public static object? TryGetActiveInstance()
    {
        EnsureWindows();
        var type = GetInventorType();
        try
        {
            return GetActiveObject(type.GUID, IntPtr.Zero);
        }
        catch (COMException)
        {
            return null;
        }
    }

    public static object CreateInstance()
    {
        EnsureWindows();
        return Activator.CreateInstance(GetInventorType())
               ?? throw new InvalidOperationException("Failed to create Inventor COM instance.");
    }

    public static object GetOrCreateInventorInstance()
    {
        return TryGetActiveInstance() ?? CreateInstance();
    }

    public static void ReleaseComObject(object? obj)
    {
        if (obj != null)
        {
            Marshal.ReleaseComObject(obj);
        }
    }
}
