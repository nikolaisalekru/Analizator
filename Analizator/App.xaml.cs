using System.Windows;

namespace Analizator;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        if (AnalysisWorker.IsWorkerInvocation(e.Args))
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            Shutdown(AnalysisWorker.Run(e.Args));
            return;
        }

        base.OnStartup(e);
        MainWindow = new MainWindow();
        MainWindow.Show();
    }
}
