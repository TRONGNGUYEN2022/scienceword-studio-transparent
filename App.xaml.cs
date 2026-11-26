using System.Windows;

namespace ScienceWordWebViewOverlay
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            var mainWin = new MainWindow();
            mainWin.Show();

            var dock = new FloatingDockWindow(mainWin);
            dock.Owner = mainWin;
            dock.Show();
        }
    }
}
