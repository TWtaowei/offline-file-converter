// 本项目自有代码及其修改版仅供非商业使用，禁止转售及企业内部业务使用。见根目录 LICENSE。
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

// A short-lived renderer using the stable prefixes of LibreOfficeKit's C structs.
// The ABI is defined in include/LibreOfficeKit/LibreOfficeKit.h in LibreOffice.
static class RenderPdf {
    [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool SetDllDirectory(string directory);
    [DllImport("mergedlo.dll", CallingConvention = CallingConvention.Cdecl)]
    static extern IntPtr libreofficekit_hook_2(byte[] directory, byte[] profileUrl);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate void Destroy(IntPtr instance);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate IntPtr DocumentLoad(IntPtr office, byte[] url);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] delegate int SaveAs(IntPtr document, byte[] url, byte[] format, IntPtr filterOptions);
    [StructLayout(LayoutKind.Sequential)] struct OfficeTable { public UIntPtr Size; public IntPtr Destroy, DocumentLoad; }
    [StructLayout(LayoutKind.Sequential)] struct DocumentTable { public UIntPtr Size; public IntPtr Destroy, SaveAs; }
    static byte[] Utf8(string value) { return Encoding.UTF8.GetBytes(value + "\0"); }
    static T Table<T>(IntPtr instance) { return (T)Marshal.PtrToStructure(Marshal.ReadIntPtr(instance), typeof(T)); }
    static T Method<T>(IntPtr address) where T : class { return (T)(object)Marshal.GetDelegateForFunctionPointer(address,typeof(T)); }

    [STAThread]
    static int Main(string[] args) {
        if (args.Length != 4) return 2;
        string engine = Path.GetFullPath(args[0]);
        if (!SetDllDirectory(engine)) return 3;
        Environment.SetEnvironmentVariable("PATH", engine + ";" + Environment.GetEnvironmentVariable("PATH"));
        Environment.SetEnvironmentVariable("SAL_DISABLESKIA", "1");
        Environment.SetEnvironmentVariable("SAL_DISABLE_OPENCL", "1");
        try {
            var office = libreofficekit_hook_2(Utf8(engine), Utf8(new Uri(args[1]).AbsoluteUri));
            if (office == IntPtr.Zero) return 4;
            var methods = Table<OfficeTable>(office);
            try {
                var document = Method<DocumentLoad>(methods.DocumentLoad)(office, Utf8(new Uri(args[2]).AbsoluteUri));
                if (document == IntPtr.Zero) return 5;
                var doc = Table<DocumentTable>(document);
                try { return Method<SaveAs>(doc.SaveAs)(document, Utf8(new Uri(args[3]).AbsoluteUri), Utf8("pdf"), IntPtr.Zero) != 0 ? 0 : 6; }
                finally { Method<Destroy>(doc.Destroy)(document); }
            } finally { Method<Destroy>(methods.Destroy)(office); }
        } catch (Exception ex) { Console.Error.WriteLine(ex.Message); return 1; }
    }
}
