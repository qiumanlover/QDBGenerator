namespace QDBGenerator
{
    using System;
    using System.CodeDom.Compiler;
    using System.Diagnostics;
    using System.Windows;

    public class App : Application
    {
        [GeneratedCode("PresentationBuildTasks", "4.0.0.0"), DebuggerNonUserCode]
        public void InitializeComponent()
        {
            base.StartupUri = new Uri("MainWindow.xaml", UriKind.Relative);
        }

        [STAThread, DebuggerNonUserCode, GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
        public static void Main()
        {
            App app = new App();
            app.InitializeComponent();
            app.Run();
        }
    }
}

