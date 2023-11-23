namespace SejinTraceability
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {        
            ApplicationConfiguration.Initialize();            
            TraceabilityForm form = new TraceabilityForm();            
            form.InitializeFormTrace();          
            Application.Run(form);
        }
    }
}