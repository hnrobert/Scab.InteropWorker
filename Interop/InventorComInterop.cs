using System.Runtime.InteropServices;

namespace Scab.InteropWorker.Interop;

public static class InventorComInterop
{
    private const string InventorProgId = "Inventor.Application";

    [DllImport("oleaut32.dll", PreserveSig = false)]
    [return: MarshalAs(UnmanagedType.IDispatch)]
    private static extern object GetActiveObject(
        [In, MarshalAs(UnmanagedType.LPStruct)] Guid clsid);

    public static object? GetOrCreateInventorInstance()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            throw new PlatformNotSupportedException("Inventor COM interop requires Windows");

        var type = Type.GetTypeFromProgID(InventorProgId)
                   ?? throw new InvalidOperationException("Inventor is not installed or not registered");

        try
        {
            return GetActiveObject(type.GUID);
        }
        catch (COMException)
        {
            return Activator.CreateInstance(type);
        }
    }

    public static void ReleaseComObject(object? obj)
    {
        if (obj != null)
        {
            Marshal.ReleaseComObject(obj);
        }
    }
}
