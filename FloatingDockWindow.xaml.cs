using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using ScienceWordWebViewOverlay.Utils;

namespace ScienceWordWebViewOverlay
{
    public partial class FloatingDockWindow : Window
    {
        private readonly MainWindow _main;
        private readonly DispatcherTimer _pollTimer;

        public FloatingDockWindow(MainWindow main)
        {
            InitializeComponent();
            _main = main;
            SourceInitialized += (_, _) => Win32Bridge.HideFromAltTab(this);

            _pollTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(450) };
            _pollTimer.Tick += (_, _) =>
            {
                var (hwnd, proc, _) = Win32Bridge.GetForegroundAppInfo();
                if (hwnd == IntPtr.Zero) return;
                var myHwnd = new WindowInteropHelper(this).Handle;
                var mainHwnd = new WindowInteropHelper(_main).Handle;
                if (hwnd == myHwnd || hwnd == mainHwnd) return;
                _main.LastTargetHwnd = hwnd;
                _main.LastTargetProcName = proc;
                TxtApp.Text = "Đích: " + (string.IsNullOrWhiteSpace(proc) ? "App khác" : proc);
            };
            _pollTimer.Start();
        }

        private void Grip_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed) DragMove();
        }

        private void BtnPass_Checked(object sender, RoutedEventArgs e) => _main.SetPassthrough(true);
        private void BtnPass_Unchecked(object sender, RoutedEventArgs e) => _main.SetPassthrough(false);

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            _pollTimer.Stop();
            Application.Current.Shutdown();
        }
    }
}
