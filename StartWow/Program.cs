using Microsoft.Win32;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace StartWow
{
    internal class Program
    {
        private static string _pathToWowExe;
        private const string keyBase = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
        private static int _x;
        private static int _y;

        [DllImport("user32.dll")]
        public static extern void SwitchToThisWindow(IntPtr hWnd, bool turnon);
        
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        static void Main(string[] args)
        {
            _pathToWowExe = !string.IsNullOrEmpty(ConfigurationManager.AppSettings["pathToWowExe"]) ?
                ConfigurationManager.AppSettings["pathToWowExe"] : GetPathForExe();

            _x = Int32.Parse(ConfigurationManager.AppSettings["x"]);
            _y = Int32.Parse(ConfigurationManager.AppSettings["y"]);

            while (true)
            {
                StartWow_Updated();
                Thread.Sleep(TimeSpan.FromMinutes(5));
                var proc = Process.GetProcesses().Where(x => x.ProcessName.ToLower().Contains("wowclassic")).First();
                proc.Kill();
            }


            //StartWow();
            //WaitForTimeToPass(
            //    TimeSpan.TryParse(ConfigurationManager.AppSettings["waitUntillRun"], out TimeSpan time) ?
            //    time : TimeSpan.Zero);
        }


        private static void TypeString(string str)
        {
            foreach (var chr in str)
            {
                Thread.Sleep(200);
                SendKeys.SendWait(chr.ToString());
            }
        }

        private static void StartWow()
        {
            SwitchToBattleNetWindow();


            AutomationElement battleNetMainWindow = null;
            int i = 0;
            while (battleNetMainWindow == null && i < 50)
            {
                battleNetMainWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Battle.net"));
                i++;
                Thread.Sleep(100);
            }
            if (battleNetMainWindow == null)
            {
                throw new InvalidOperationException("Battle.net isn't running");
            }

            var wowClassicTab = battleNetMainWindow.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "World of Warcraft Classic"));
            Click(wowClassicTab);
            CachePropertiesWithScope(battleNetMainWindow);
        }

        private static void StartWow_Updated()
        {
            SwitchToBattleNetWindow();
            Thread.Sleep(5000);
            MouseOperations.SetCursorPosition(_x, _y);
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
            MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
            Thread.Sleep(15000);
            var proc = Process.GetProcesses().Where(x => x.ProcessName.ToLower().Contains("wowclassic")).First();
            ShowWindow(proc.MainWindowHandle, 2);
            SwitchToBattleNetWindow();

        }

        private static void SwitchToBattleNetWindow()
        {
            Process[] processes = Process.GetProcessesByName("Battle.net");
            foreach (var proc in processes)
            {
                SwitchToThisWindow(proc.MainWindowHandle, true);
                ShowWindow(proc.MainWindowHandle, 1);
                ShowWindow(proc.MainWindowHandle, 3);

            }
        }

        private static void SetUpTimer(TimeSpan alertTime)
        {
            TimeSpan timeToGo = alertTime - DateTime.Now.TimeOfDay;

            while (true)
            {
                if (timeToGo < TimeSpan.Zero)
                {
                    StartWow();
                    break;
                }
                else
                {
                    Thread.Sleep(1000);
                    timeToGo = alertTime - DateTime.Now.TimeOfDay;

                }
            }
        }

        private static void WaitForTimeToPass(TimeSpan waitingTime)
        {
            Stopwatch sw = Stopwatch.StartNew();
            double waitingMiliseconds = waitingTime.TotalMilliseconds;

            while (sw.ElapsedMilliseconds < waitingMiliseconds)
                Thread.Sleep(1000);

            StartWow();
        }

        private static void CachePropertiesWithScope(AutomationElement elementMain)
        {
            AutomationElement playButton = null;
            CacheRequest c = new CacheRequest();
            c.Add(AutomationElement.NameProperty);
            c.Add(SelectionItemPattern.Pattern);
            c.Add(SelectionItemPattern.SelectionContainerProperty);
            c.TreeScope = TreeScope.Element | TreeScope.Children;
            using (c.Activate())
            {
                //var a = elementMain.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Play: Wrath of the Lich King Classic, Version: 3.4.0.45435"));

                playButton = elementMain.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, $"Play: Wrath of the Lich King Classic, Version: {GetWowVerstion()}"));
                Click(playButton);
            }
        }

        private static string GetPathForExe()
        {
            RegistryKey localMachine = Registry.LocalMachine;
            RegistryKey fileKey = localMachine.OpenSubKey(keyBase);

            foreach (var subKey in fileKey.GetSubKeyNames())
            {
                var a = fileKey.OpenSubKey(subKey);
                if (a.ValueCount != 0)
                {
                    var b = a.GetValueNames();
                    foreach (var name in b)
                    {
                        if (a.GetValue(name).ToString().Contains("WoWClassic.exe"))
                            return a.GetValue(name).ToString();
                    }
                }
            }
            throw new Exception("Can't find path to WoWClassic.exe, provide it manually in App.config file");
        }

        private static string GetWowVerstion()
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(_pathToWowExe);
            return versionInfo.FileVersion;
        }

        private static void Click(AutomationElement element)
        {
            var invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();
        }
    }
}
