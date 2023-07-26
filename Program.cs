using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace MicroscopeConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string st = args[0];
            st = st.Replace('_',' ');
            Microscope.Initialize(st);
            do
            {
                string s = Console.ReadLine();
                if(s.Length > 0 && s.StartsWith("Command:"))
                {
                    s = s.Replace("Command:", "");
                    CommandRunner.Command com = JsonConvert.DeserializeObject<CommandRunner.Command> (s);
                    CommandRunner.Run(com);
                }
                System.Threading.Thread.Sleep(20);
            } while (true);
        }
    }
}
