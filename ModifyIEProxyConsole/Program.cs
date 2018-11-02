using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ModifyIEProxyConsole.Controller;

namespace ModifyIEProxyConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread th = new Thread(ThreadMain);
            th.IsBackground = true;
            th.Start();
            Console.Read();
        }

        private static void ThreadMain()
        {
            int threadSleepTime = 10000;

            bool enabled = true;
            bool global = false;
            int localPort = 1080;
            string localAuthPassword = "B3kN-zAKJB4uQ0Qo4RvT";

            while (1==1)
            {
                SystemProxy.SetFixedProxyInfo(enabled, global, localPort, localAuthPassword);
                Thread.Sleep(threadSleepTime);
            }
        }
    }
}
