using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Interop;

namespace ScienceWordWebViewOverlay.Utils
{
    public class OfficeTargetInfo
    {
        public string AppKind { get; set; } = "";
        public double DipPerPt { get; set; } = 96.0 / 72.0;
        public double RefDipLeft { get; set; }
        public double RefDipTop { get; set; }
        public double RefDocLeftPt { get; set; }
        public double RefDocTopPt { get; set; }
        public int RefRelH { get; set; } = 1;
        public int RefRelV { get; set; } = 1;
        public object? RefAnchor { get; set; }
        public object? SourceShape { get; set; }
    }

    public static class Win32Bridge
    {
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        private const byte VK_CONTROL = 0x11;
        private const byte VK_V = 0x56;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const int MsoTrue = -1;
        public const string AltPrefix = "SW_WV2_V1|";

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hwnd, int index);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public static void SetClickThrough(Window window, bool enable)
        {
            var hwnd = new WindowInteropHelper(window).Handle;
            if (hwnd == IntPtr.Zero) return;
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            if (enable) exStyle |= WS_EX_TRANSPARENT;
            else exStyle &= ~WS_EX_TRANSPARENT;
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);
        }

        public static void HideFromAltTab(Window window)
        {
            var hwnd = new WindowInteropHelper(window).Handle;
            if (hwnd == IntPtr.Zero) return;
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            exStyle |= WS_EX_TOOLWINDOW;
            SetWindowLong(hwnd, GWL_EXSTYLE, exStyle);
        }

        public static (IntPtr hwnd, string procName, string title) GetForegroundAppInfo()
        {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return (IntPtr.Zero, "", "");
            string proc = "";
            try
            {
                GetWindowThreadProcessId(hwnd, out uint pid);
                using var p = Process.GetProcessById((int)pid);
                proc = p.ProcessName;
            }
            catch { }
            var sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            return (hwnd, proc, sb.ToString());
        }

        public static bool SendPasteToWindow(IntPtr targetHwnd)
        {
            if (targetHwnd == IntPtr.Zero || !IsWindow(targetHwnd)) return false;
            try
            {
                SetForegroundWindow(targetHwnd);
                Thread.Sleep(120);
                keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
                keybd_event(VK_V, 0, 0, UIntPtr.Zero);
                keybd_event(VK_V, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                return true;
            }
            catch { return false; }
        }

        private static object? GetActiveComObject(string progId)
        {
            if (CLSIDFromProgID(progId, out Guid clsid) < 0) return null;
            return GetActiveObject(ref clsid, IntPtr.Zero, out object obj) >= 0 ? obj : null;
        }

        public static OfficeTargetInfo? CaptureOfficeSelectionContext(string procName, double dpiScaleX, double dpiScaleY)
        {
            string p = procName.ToUpperInvariant();
            try
            {
                if (p == "WINWORD")
                {
                    object? wObj = GetActiveComObject("Word.Application");
                    if (wObj == null) return null;
                    dynamic wApp = wObj;
                    if ((int)wApp.Documents.Count == 0) return null;
                    object winObj = wApp.ActiveWindow;
                    dynamic? sr = null;
                    try { sr = wApp.Selection.ShapeRange; } catch { }
                    if (sr == null || (int)sr.Count == 0) return null;

                    object shObj = sr.Item(1);
                    dynamic sh = shObj;
                    double zoomPct = 100.0;
                    try { zoomPct = Convert.ToDouble(wApp.ActivePane.View.Zoom.Percentage); } catch { }

                    Rect pxRect = TryWordGetPoint(winObj, sr);
                    double dipPerPt = (96.0 / 72.0) * (zoomPct / 100.0);
                    double wPt = Math.Max(1.0, Convert.ToDouble(sh.Width));
                    if (!pxRect.IsEmpty && pxRect.Width > 4) dipPerPt = (pxRect.Width / dpiScaleX) / wPt;

                    var info = new OfficeTargetInfo
                    {
                        AppKind = "WINWORD",
                        DipPerPt = dipPerPt,
                        RefDipLeft = !pxRect.IsEmpty ? pxRect.Left / dpiScaleX : Convert.ToDouble(sh.Left) * dipPerPt,
                        RefDipTop = !pxRect.IsEmpty ? pxRect.Top / dpiScaleY : Convert.ToDouble(sh.Top) * dipPerPt,
                        RefDocLeftPt = Convert.ToDouble(sh.Left),
                        RefDocTopPt = Convert.ToDouble(sh.Top)
                    };
                    try { info.RefRelH = (int)sh.RelativeHorizontalPosition; } catch { }
                    try { info.RefRelV = (int)sh.RelativeVerticalPosition; } catch { }
                    try { info.RefAnchor = sh.Anchor; } catch { }
                    return info;
                }
                if (p == "POWERPNT")
                {
                    object? pObj = GetActiveComObject("PowerPoint.Application");
                    if (pObj == null) return null;
                    dynamic pApp = pObj;
                    if ((int)pApp.Presentations.Count == 0) return null;
                    dynamic win = pApp.ActiveWindow;
                    dynamic? sr = null;
                    try { sr = win.Selection.ShapeRange; } catch { }
                    if (sr == null || (int)sr.Count == 0) return null;

                    dynamic sh = sr.Item(1);
                    float lPt = Convert.ToSingle(sh.Left), tPt = Convert.ToSingle(sh.Top);
                    float wPt = Math.Max(1f, Convert.ToSingle(sh.Width)), hPt = Math.Max(1f, Convert.ToSingle(sh.Height));
                    double x1 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt));
                    double y1 = Convert.ToDouble(win.PointsToScreenPixelsY(tPt));
                    double x2 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt + wPt));
                    double dipPerPt = Math.Max(0.2, ((x2 - x1) / dpiScaleX) / wPt);

                    return new OfficeTargetInfo
                    {
                        AppKind = "POWERPNT",
                        DipPerPt = dipPerPt,
                        RefDipLeft = x1 / dpiScaleX,
                        RefDipTop = y1 / dpiScaleY,
                        RefDocLeftPt = lPt,
                        RefDocTopPt = tPt
                    };
                }
                if (p == "EXCEL")
                {
                    object? xObj = GetActiveComObject("Excel.Application");
                    if (xObj == null) return null;
                    dynamic xApp = xObj;
                    if ((int)xApp.Workbooks.Count == 0) return null;
                    dynamic win = xApp.ActiveWindow;
                    dynamic? sr = null;
                    try { sr = xApp.Selection.ShapeRange; } catch { }
                    if (sr == null || (int)sr.Count == 0) return null;

                    dynamic sh = sr.Item(1);
                    double lPt = Convert.ToDouble(sh.Left), tPt = Convert.ToDouble(sh.Top);
                    double wPt = Math.Max(1.0, Convert.ToDouble(sh.Width));
                    double x1 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt));
                    double y1 = Convert.ToDouble(win.PointsToScreenPixelsY(tPt));
                    double x2 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt + wPt));
                    double dipPerPt = Math.Max(0.2, ((x2 - x1) / dpiScaleX) / wPt);

                    return new OfficeTargetInfo
                    {
                        AppKind = "EXCEL",
                        DipPerPt = dipPerPt,
                        RefDipLeft = x1 / dpiScaleX,
                        RefDipTop = y1 / dpiScaleY,
                        RefDocLeftPt = lPt,
                        RefDocTopPt = tPt
                    };
                }
            }
            catch { }
            return null;
        }

        public static bool InsertPngAtScreenRect(OfficeTargetInfo? ctx, string procName, string pngPath, Rect screenRectDip, string? embeddedAltText, out string? error)
        {
            error = null;
            string kind = ctx?.AppKind ?? procName.ToUpperInvariant();
            try
            {
                if (kind == "WINWORD")
                {
                    object? wObj = GetActiveComObject("Word.Application");
                    if (wObj == null) { error = "Không tìm thấy Word."; return false; }
                    dynamic wApp = wObj;
                    dynamic doc = wApp.ActiveDocument;

                    double dipPerPt = (ctx != null && ctx.DipPerPt > 0.1) ? ctx.DipPerPt : (96.0 / 72.0);
                    double leftPt = ctx != null ? ctx.RefDocLeftPt + (screenRectDip.Left - ctx.RefDipLeft) / dipPerPt : 72.0;
                    double topPt = ctx != null ? ctx.RefDocTopPt + (screenRectDip.Top - ctx.RefDipTop) / dipPerPt : 72.0;
                    double wPt = screenRectDip.Width / dipPerPt;
                    double hPt = screenRectDip.Height / dipPerPt;

                    dynamic pic = ctx?.RefAnchor != null
                        ? doc.Shapes.AddPicture(pngPath, false, true, (float)leftPt, (float)topPt, (float)wPt, (float)hPt, ctx.RefAnchor)
                        : doc.Shapes.AddPicture(pngPath, false, true, (float)leftPt, (float)topPt, (float)wPt, (float)hPt);

                    try { pic.LockAspectRatio = MsoTrue; } catch { }
                    try { pic.WrapFormat.Type = 3; } catch { }
                    if (ctx != null)
                    {
                        try { pic.RelativeHorizontalPosition = ctx.RefRelH; } catch { }
                        try { pic.RelativeVerticalPosition = ctx.RefRelV; } catch { }
                    }
                    pic.Left = (float)leftPt; pic.Top = (float)topPt;
                    pic.Width = (float)wPt; pic.Height = (float)hPt;
                    if (!string.IsNullOrEmpty(embeddedAltText))
                    {
                        try { pic.AlternativeText = embeddedAltText; } catch { }
                    }
                    if (ctx?.SourceShape != null)
                    {
                        try { ((dynamic)ctx.SourceShape).Delete(); } catch { }
                        ctx.SourceShape = pic;
                    }
                    return true;
                }
                if (kind == "POWERPNT")
                {
                    object? pObj = GetActiveComObject("PowerPoint.Application");
                    if (pObj == null) { error = "Không tìm thấy PowerPoint."; return false; }
                    dynamic pApp = pObj;
                    dynamic slide = pApp.ActiveWindow.View.Slide;

                    double dipPerPt = (ctx != null && ctx.DipPerPt > 0.1) ? ctx.DipPerPt : (96.0 / 72.0);
                    double leftPt = ctx != null ? ctx.RefDocLeftPt + (screenRectDip.Left - ctx.RefDipLeft) / dipPerPt : 60.0;
                    double topPt = ctx != null ? ctx.RefDocTopPt + (screenRectDip.Top - ctx.RefDipTop) / dipPerPt : 60.0;
                    double wPt = screenRectDip.Width / dipPerPt;
                    double hPt = screenRectDip.Height / dipPerPt;

                    dynamic pic = slide.Shapes.AddPicture(pngPath, 0, -1, (float)leftPt, (float)topPt, (float)wPt, (float)hPt);
                    if (!string.IsNullOrEmpty(embeddedAltText))
                    {
                        try { pic.AlternativeText = embeddedAltText; } catch { }
                    }
                    if (ctx?.SourceShape != null)
                    {
                        try { ((dynamic)ctx.SourceShape).Delete(); } catch { }
                        ctx.SourceShape = pic;
                    }
                    return true;
                }
                if (kind == "EXCEL")
                {
                    object? xObj = GetActiveComObject("Excel.Application");
                    if (xObj == null) { error = "Không tìm thấy Excel."; return false; }
                    dynamic xApp = xObj;
                    dynamic sheet = xApp.ActiveSheet;

                    double dipPerPt = (ctx != null && ctx.DipPerPt > 0.1) ? ctx.DipPerPt : (96.0 / 72.0);
                    double leftPt = ctx != null ? ctx.RefDocLeftPt + (screenRectDip.Left - ctx.RefDipLeft) / dipPerPt : 60.0;
                    double topPt = ctx != null ? ctx.RefDocTopPt + (screenRectDip.Top - ctx.RefDipTop) / dipPerPt : 60.0;
                    double wPt = screenRectDip.Width / dipPerPt;
                    double hPt = screenRectDip.Height / dipPerPt;

                    dynamic pic = sheet.Shapes.AddPicture(pngPath, 0, -1, (float)leftPt, (float)topPt, (float)wPt, (float)hPt);
                    if (!string.IsNullOrEmpty(embeddedAltText))
                    {
                        try { pic.AlternativeText = embeddedAltText; } catch { }
                    }
                    if (ctx?.SourceShape != null)
                    {
                        try { ((dynamic)ctx.SourceShape).Delete(); } catch { }
                        ctx.SourceShape = pic;
                    }
                    return true;
                }
            }
            catch (Exception ex) { error = ex.Message; return false; }

            error = "Không phải ứng dụng Office.";
            return false;
        }

        public static (string? projectJson, OfficeTargetInfo? ctx)? TryReadEmbeddedProjectFromOffice(string procName, double dpiScaleX, double dpiScaleY)
        {
            string p = procName.ToUpperInvariant();
            try
            {
                if (p == "WINWORD")
                {
                    object? wObj = GetActiveComObject("Word.Application");
                    if (wObj == null) return null;
                    dynamic wApp = wObj;
                    dynamic? sr = null;
                    try { sr = wApp.Selection.ShapeRange; } catch { return null; }
                    if (sr == null || (int)sr.Count != 1) return null;
                    object shObj = sr.Item(1);
                    dynamic sh = shObj;
                    string alt = "";
                    try { alt = Convert.ToString(sh.AlternativeText) ?? ""; } catch { return null; }
                    if (!alt.StartsWith(AltPrefix, StringComparison.Ordinal)) return null;

                    string json = Encoding.UTF8.GetString(Convert.FromBase64String(alt.Substring(AltPrefix.Length)));
                    Rect pxRect = TryWordGetPoint(wApp.ActiveWindow, sr);
                    double dipPerPt = (96.0 / 72.0);
                    double wPt = Math.Max(1.0, Convert.ToDouble(sh.Width));
                    if (!pxRect.IsEmpty && pxRect.Width > 4) dipPerPt = (pxRect.Width / dpiScaleX) / wPt;

                    var info = new OfficeTargetInfo
                    {
                        AppKind = "WINWORD",
                        DipPerPt = dipPerPt,
                        RefDipLeft = !pxRect.IsEmpty ? pxRect.Left / dpiScaleX : Convert.ToDouble(sh.Left) * dipPerPt,
                        RefDipTop = !pxRect.IsEmpty ? pxRect.Top / dpiScaleY : Convert.ToDouble(sh.Top) * dipPerPt,
                        RefDocLeftPt = Convert.ToDouble(sh.Left),
                        RefDocTopPt = Convert.ToDouble(sh.Top),
                        SourceShape = shObj
                    };
                    try { info.RefRelH = (int)sh.RelativeHorizontalPosition; } catch { }
                    try { info.RefRelV = (int)sh.RelativeVerticalPosition; } catch { }
                    try { info.RefAnchor = sh.Anchor; } catch { }
                    return (json, info);
                }
                if (p == "POWERPNT")
                {
                    object? pObj = GetActiveComObject("PowerPoint.Application");
                    if (pObj == null) return null;
                    dynamic pApp = pObj;
                    dynamic win = pApp.ActiveWindow;
                    dynamic? sr = null;
                    try { sr = win.Selection.ShapeRange; } catch { return null; }
                    if (sr == null || (int)sr.Count != 1) return null;
                    object shObj = sr.Item(1);
                    dynamic sh = shObj;
                    string alt = "";
                    try { alt = Convert.ToString(sh.AlternativeText) ?? ""; } catch { return null; }
                    if (!alt.StartsWith(AltPrefix, StringComparison.Ordinal)) return null;

                    string json = Encoding.UTF8.GetString(Convert.FromBase64String(alt.Substring(AltPrefix.Length)));
                    float lPt = Convert.ToSingle(sh.Left), tPt = Convert.ToSingle(sh.Top), wPt = Math.Max(1f, Convert.ToSingle(sh.Width));
                    double x1 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt));
                    double y1 = Convert.ToDouble(win.PointsToScreenPixelsY(tPt));
                    double x2 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt + wPt));
                    var info = new OfficeTargetInfo
                    {
                        AppKind = "POWERPNT",
                        DipPerPt = Math.Max(0.2, ((x2 - x1) / dpiScaleX) / wPt),
                        RefDipLeft = x1 / dpiScaleX,
                        RefDipTop = y1 / dpiScaleY,
                        RefDocLeftPt = lPt,
                        RefDocTopPt = tPt,
                        SourceShape = shObj
                    };
                    return (json, info);
                }
                if (p == "EXCEL")
                {
                    object? xObj = GetActiveComObject("Excel.Application");
                    if (xObj == null) return null;
                    dynamic xApp = xObj;
                    dynamic win = xApp.ActiveWindow;
                    dynamic? sr = null;
                    try { sr = xApp.Selection.ShapeRange; } catch { return null; }
                    if (sr == null || (int)sr.Count != 1) return null;
                    object shObj = sr.Item(1);
                    dynamic sh = shObj;
                    string alt = "";
                    try { alt = Convert.ToString(sh.AlternativeText) ?? ""; } catch { return null; }
                    if (!alt.StartsWith(AltPrefix, StringComparison.Ordinal)) return null;

                    string json = Encoding.UTF8.GetString(Convert.FromBase64String(alt.Substring(AltPrefix.Length)));
                    double lPt = Convert.ToDouble(sh.Left), tPt = Convert.ToDouble(sh.Top), wPt = Math.Max(1.0, Convert.ToDouble(sh.Width));
                    double x1 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt));
                    double y1 = Convert.ToDouble(win.PointsToScreenPixelsY(tPt));
                    double x2 = Convert.ToDouble(win.PointsToScreenPixelsX(lPt + wPt));
                    var info = new OfficeTargetInfo
                    {
                        AppKind = "EXCEL",
                        DipPerPt = Math.Max(0.2, ((x2 - x1) / dpiScaleX) / wPt),
                        RefDipLeft = x1 / dpiScaleX,
                        RefDipTop = y1 / dpiScaleY,
                        RefDocLeftPt = lPt,
                        RefDocTopPt = tPt,
                        SourceShape = shObj
                    };
                    return (json, info);
                }
            }
            catch { }
            return null;
        }

        private static Rect TryWordGetPoint(object activeWindow, object targetObj)
        {
            try
            {
                object[] args = { 0, 0, 0, 0, targetObj };
                var pm = new ParameterModifier(5);
                pm[0] = true; pm[1] = true; pm[2] = true; pm[3] = true; pm[4] = false;
                activeWindow.GetType().InvokeMember("GetPoint", BindingFlags.InvokeMethod, null, activeWindow, args, new[] { pm }, null, null);
                return new Rect(Convert.ToInt32(args[0]), Convert.ToInt32(args[1]),
                    Math.Max(0, Convert.ToInt32(args[2])), Math.Max(0, Convert.ToInt32(args[3])));
            }
            catch { return Rect.Empty; }
        }
    }
}
