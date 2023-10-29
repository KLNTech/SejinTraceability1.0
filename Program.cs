namespace SejinTraceability
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {        
            ApplicationConfiguration.Initialize();            
            straceabilitysystem form = new straceabilitysystem();            
            form.InitializeFormTrace();          
            Application.Run(form);
        }
    }
}