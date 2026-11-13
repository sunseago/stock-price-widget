// =====================================================================
//  PDD Stock Widget - single-file EXE launcher
//  Embeds PDD-Widget.ps1 + PDD-Fetch.ps1 as resources, extracts them to
//  %LOCALAPPDATA%\PDDStockWidget on every run, then launches the widget
//  hidden via Windows PowerShell (built into every Windows machine).
//
//  Usage:  PDD股价悬浮窗.exe            -> start the widget
//          PDD股价悬浮窗.exe -shortcut  -> also ensure a Desktop shortcut
// =====================================================================
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

internal static class Launcher
{
    private static string BaseDir
    {
        get
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PDDStockWidget");
        }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            string dir = BaseDir;
            Directory.CreateDirectory(dir);

            Extract("PDD-Widget.ps1", dir);
            Extract("PDD-Fetch.ps1", dir);

            bool wantShortcut = false;
            if (args != null)
            {
                foreach (string a in args)
                {
                    if (string.Equals(a, "-shortcut", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(a, "/shortcut", StringComparison.OrdinalIgnoreCase))
                    {
                        wantShortcut = true;
                    }
                }
            }
            if (wantShortcut) { TryCreateDesktopShortcut(); }

            string ps = Path.Combine(dir, "PDD-Widget.ps1");
            ProcessStartInfo psi = new ProcessStartInfo("powershell.exe");
            psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File \"" + ps + "\"";
            psi.CreateNoWindow = true;
            psi.UseShellExecute = false;
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            // 让脚本知道自身由哪个 EXE 启动, 供"开机自启动"创建正确的快捷方式
            psi.EnvironmentVariables["PDDWIDGET_EXE"] = Assembly.GetExecutingAssembly().Location;
            using (Process p = Process.Start(psi)) { }
            return 0;
        }
        catch (Exception ex)
        {
            TryLog(ex.ToString());
            return 1;
        }
    }

    private static void Extract(string resourceName, string dir)
    {
        string target = Path.Combine(dir, resourceName);
        using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
        {
            if (s == null)
            {
                throw new InvalidOperationException("Missing embedded resource: " + resourceName);
            }
            using (FileStream fs = new FileStream(target, FileMode.Create, FileAccess.Write, FileShare.Read))
            {
                byte[] buf = new byte[65536];
                int n;
                while ((n = s.Read(buf, 0, buf.Length)) > 0) { fs.Write(buf, 0, n); }
            }
        }
    }

    private static void TryCreateDesktopShortcut()
    {
        try
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string lnk = Path.Combine(desktop, "股价悬浮窗.lnk");
            if (File.Exists(lnk)) { return; }
            string exe = Assembly.GetExecutingAssembly().Location;

            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null) { return; }
            object wsh = Activator.CreateInstance(shellType);
            object sc = shellType.InvokeMember("CreateShortcut",
                BindingFlags.InvokeMethod, null, wsh, new object[] { lnk });
            Type scType = sc.GetType();
            scType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, sc, new object[] { exe });
            scType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, sc,
                new object[] { Path.GetDirectoryName(exe) });
            scType.InvokeMember("Description", BindingFlags.SetProperty, null, sc,
                new object[] { "股价悬浮窗" });
            scType.InvokeMember("Save", BindingFlags.InvokeMethod, null, sc, null);
        }
        catch { }
    }

    private static void TryLog(string text)
    {
        try
        {
            File.AppendAllText(Path.Combine(Path.GetTempPath(), "pdd-widget-launch.log"),
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + text + Environment.NewLine);
        }
        catch { }
    }
}
