namespace AudioSwitcherService
{
    using System;
    using System.IO;
    using System.Management;
    using System.ServiceProcess;

    using VisioForge.Libs.NAudio.CoreAudioApi;

    public partial class AudioSwitcherService : ServiceBase
    {
        private string _defaultCommunicationsAudio;

        private string _defaultConsoleAudio;

        private string _defaultMultimediaAudio;

        private ManagementEventWatcher _startWatch;

        private ManagementEventWatcher _stopWatch;

        private string PROCESS_NAME = null;

        public AudioSwitcherService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            string configFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\audioSwitcher.txt";
            if (!File.Exists(configFile))
            {
                File.AppendAllText(configFile, "replaceMe.exe");
            }

            PROCESS_NAME = File.ReadAllLines(configFile)[0];

            WaitForProcess();
            if (Environment.UserInteractive)
            {
                Console.WriteLine("I have been started!");
            }
        }

        protected override void OnStop()
        {
            _startWatch.EventArrived -= startWatch_EventArrived;
            _startWatch.Stop();

            _stopWatch.EventArrived -= stopWatch_EventArrived;
            _stopWatch.Stop();

            if (Environment.UserInteractive)
            {
                Console.WriteLine("I have been stopped!");
            }
        }

        void GetDefaultAudioDevices()
        {
            var enumerator = new MMDeviceEnumerator();
            var deviceCommunications = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Communications);
            _defaultCommunicationsAudio = deviceCommunications.ID;

            var deviceConsole = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Console);
            _defaultConsoleAudio = deviceConsole.ID;

            var deviceMultimedia = enumerator.GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
            _defaultMultimediaAudio = deviceMultimedia.ID;
        }

        void startWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            try
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine("Process started: {0}", e.NewEvent.Properties["ProcessName"].Value);
                }

                string processName = (string) e.NewEvent.Properties["ProcessName"].Value;
                if (processName.StartsWith(PROCESS_NAME))
                {
                    if (_defaultCommunicationsAudio == null && _defaultConsoleAudio == null && _defaultMultimediaAudio == null)
                    {
                        GetDefaultAudioDevices();

                        if (_defaultCommunicationsAudio != null)
                        {
                            CoreAudioApi.PolicyConfigClient client = new CoreAudioApi.PolicyConfigClient();
                            client.SetDefaultEndpoint(_defaultCommunicationsAudio, CoreAudioApi.ERole.eCommunications);
                            client.SetDefaultEndpoint(_defaultCommunicationsAudio, CoreAudioApi.ERole.eConsole);
                            client.SetDefaultEndpoint(_defaultCommunicationsAudio, CoreAudioApi.ERole.eMultimedia);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }

        void stopWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            try
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine("Process ended: {0}", e.NewEvent.Properties["ProcessName"].Value);
                }

                string processName = (string) e.NewEvent.Properties["ProcessName"].Value;
                if (processName.StartsWith(PROCESS_NAME))
                {
                    CoreAudioApi.PolicyConfigClient client = new CoreAudioApi.PolicyConfigClient();

                    if (_defaultCommunicationsAudio != null)
                    {
                        client.SetDefaultEndpoint(_defaultCommunicationsAudio, CoreAudioApi.ERole.eCommunications);
                    }

                    if (_defaultConsoleAudio != null)
                    {
                        client.SetDefaultEndpoint(_defaultConsoleAudio, CoreAudioApi.ERole.eConsole);
                    }

                    if (_defaultMultimediaAudio != null)
                    {
                        client.SetDefaultEndpoint(_defaultMultimediaAudio, CoreAudioApi.ERole.eMultimedia);
                    }

                    _defaultMultimediaAudio = null;
                    _defaultConsoleAudio = null;
                    _defaultCommunicationsAudio = null;
                }
            }
            catch (Exception ex)
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }

        void WaitForProcess()
        {
            _startWatch = new ManagementEventWatcher(
                new WqlEventQuery("SELECT * FROM Win32_ProcessStartTrace"));
            _startWatch.EventArrived
                += startWatch_EventArrived;
            _startWatch.Start();

            _stopWatch = new ManagementEventWatcher(
                new WqlEventQuery("SELECT * FROM Win32_ProcessStopTrace"));
            _stopWatch.EventArrived
                += stopWatch_EventArrived;
            _stopWatch.Start();
        }
    }
}