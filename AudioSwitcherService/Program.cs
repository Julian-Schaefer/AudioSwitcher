namespace AudioSwitcherService
{
    using System;
    using System.ServiceProcess;

    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            if (Environment.UserInteractive)
            {
                AudioSwitcherService service1 = new AudioSwitcherService();
                service1.TestStartupAndStop(null);
            }
            else
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new AudioSwitcherService()
                };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}