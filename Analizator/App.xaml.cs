using System.Windows;

namespace Analizator;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        if (PdfPhotoWorker.IsWorkerInvocation(e.Args))
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            Shutdown(PdfPhotoWorker.Run(e.Args));
            return;
        }

        if (MonitoringWorker.IsWorkerInvocation(e.Args))
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            Shutdown(MonitoringWorker.Run(e.Args));
            return;
        }

        if (StatisticsWorker.IsWorkerInvocation(e.Args))
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            Shutdown(StatisticsWorker.Run(e.Args));
            return;
        }

        if (DatabaseImportWorker.IsWorkerInvocation(e.Args))
        {
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            Shutdown(DatabaseImportWorker.Run(e.Args));
            return;
        }

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
