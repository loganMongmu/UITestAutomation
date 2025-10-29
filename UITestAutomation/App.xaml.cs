using System;
using System.Runtime.InteropServices;
using System.Windows;

namespace MouseMacroWpf
{
    public partial class App : Application
    {
        [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr value);
        static readonly IntPtr PMv2 = new IntPtr(-4); // PER_MONITOR_AWARE_V2

        protected override void OnStartup(StartupEventArgs e)
        {
            try { SetProcessDpiAwarenessContext(PMv2); } catch { /* ignore */ }
            base.OnStartup(e);
        }
    }
}
