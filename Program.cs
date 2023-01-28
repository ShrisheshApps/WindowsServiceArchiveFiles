using System.ServiceProcess;

namespace WindowsServiceArchiveFiles
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new ArchiveService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
