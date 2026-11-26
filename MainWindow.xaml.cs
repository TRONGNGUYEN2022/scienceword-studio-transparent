using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Web.WebView2.Core;
using ScienceWordWebViewOverlay.Utils;

namespace ScienceWordWebViewOverlay
{
    public partial class MainWindow : Window
    {
        public IntPtr LastTargetHwnd { get; set; } = IntPtr.Zero;
        public string LastTargetProcName { get; set; } = "";
        public OfficeTargetInfo? ActiveOfficeCtx { get; set; }
        public bool IsPassthrough { get; private set; } = false;

        public MainWindow()
        {
            InitializeComponent();
            SourceInitialized += (_, _) => Win32Bridge.HideFromAltTab(this);
            Loaded += OnMainWindowLoaded;
        }

        private async void OnMainWindowLoaded(object sender, RoutedEventArgs e)
        {
            await WebCanvas.EnsureCoreWebView2Async();
            WebCanvas.CoreWebView2.Settings.AreDevToolsEnabled = true;
            WebCanvas.CoreWebView2.Settings.IsStatusBarEnabled = false;
            WebCanvas.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

            string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "ScienceWordStudio.html");
            if (File.Exists(htmlPath))
                WebCanvas.CoreWebView2.Navigate(new Uri(htmlPath).AbsoluteUri);
            else
                WebCanvas.CoreWebView2.NavigateToString("<h2>Không tìm thấy file Assets/ScienceWordStudio.html</h2>");

            WebCanvas.CoreWebView2.NavigationCompleted += async (_, _) =>
            {
                await WebCanvas.CoreWebView2.ExecuteScriptAsync(
                    "if(typeof window.setOverlayTransparentMode === 'function') { window.setOverlayTransparentMode(true); const b = document.getElementById('btnWebViewTransparentToggle'); if(b){ b.classList.remove('bg-fuchsia-500','text-white'); b.classList.add('bg-amber-400','text-slate-950'); } }");
            };
        }

        public void SetPassthrough(bool enable)
        {
            IsPassthrough = enable;
            Win32Bridge.SetClickThrough(this, enable);
            if (!enable)
            {
                var dpi = VisualTreeHelper.GetDpi(this);
                ActiveOfficeCtx = Win32Bridge.CaptureOfficeSelectionContext(LastTargetProcName, dpi.DpiScaleX, dpi.DpiScaleY);
            }
        }

        public async void TryRestoreFromDoubleClick(string procName)
        {
            var dpi = VisualTreeHelper.GetDpi(this);
            var res = Win32Bridge.TryReadEmbeddedProjectFromOffice(procName, dpi.DpiScaleX, dpi.DpiScaleY);
            if (res == null || string.IsNullOrEmpty(res.Value.projectJson)) return;

            ActiveOfficeCtx = res.Value.ctx;
            SetPassthrough(false);

            string escaped = JsonSerializer.Serialize(res.Value.projectJson);
            await WebCanvas.CoreWebView2.ExecuteScriptAsync(
                $"window.postMessage({{ type: 'restore_project', projectJson: {escaped} }}, '*');");
        }

        private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                using var doc = JsonDocument.Parse(e.WebMessageAsJson);
                var root = doc.RootElement;
                string type = root.TryGetProperty("type", out var tProp) ? tProp.GetString() ?? "" : "";
                if (type != "insert_to_active_app") return;

                string projectJson = root.TryGetProperty("projectJson", out var pjProp) ? pjProp.GetString() ?? "" : "";
                var rEl = root.GetProperty("screenRect");
                double left = rEl.GetProperty("left").GetDouble();
                double top = rEl.GetProperty("top").GetDouble();
                double width = Math.Max(20, rEl.GetProperty("width").GetDouble());
                double height = Math.Max(20, rEl.GetProperty("height").GetDouble());

                string tempPng = Path.Combine(Path.GetTempPath(), $"sw_wv2_{DateTime.Now.Ticks}.png");
                var screenRectDip = new Rect(left, top, width, height);
                string? altPayload = !string.IsNullOrEmpty(projectJson)
                    ? Win32Bridge.AltPrefix + Convert.ToBase64String(Encoding.UTF8.GetBytes(projectJson))
                    : null;

                var rtb = new RenderTargetBitmap(Math.Max(16, (int)(width * 2)), Math.Max(16, (int)(height * 2)), 192, 192, PixelFormats.Pbgra32);
                var dv = new DrawingVisual();
                using (var dc = dv.RenderOpen())
                {
                    var br = new SolidColorBrush(Color.FromArgb(35, 37, 99, 235));
                    dc.DrawRectangle(br, new Pen(Brushes.DodgerBlue, 2), new Rect(0, 0, rtb.PixelWidth * 0.5, rtb.PixelHeight * 0.5));
                }
                rtb.Render(dv);
                var enc = new PngBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(rtb));
                using (var fs = File.Create(tempPng)) enc.Save(fs);

                string pUp = LastTargetProcName.ToUpperInvariant();
                if (ActiveOfficeCtx != null || pUp is "WINWORD" or "POWERPNT" or "EXCEL")
                {
                    if (Win32Bridge.InsertPngAtScreenRect(ActiveOfficeCtx, LastTargetProcName, tempPng, screenRectDip, altPayload, out _))
                        return;
                }

                Clipboard.SetImage(rtb);
                Win32Bridge.SendPasteToWindow(LastTargetHwnd);
            }
            catch { }
        }
    }
}
