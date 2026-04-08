namespace ClipSidekick;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        var monitor = new ClipboardMonitor();
        var window = new MainWindow(monitor);
        window.Create();
        window.Run();
        Environment.Exit(0);
    }
}
