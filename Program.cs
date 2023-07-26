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
            Microscope.Initialize(args[0]);
            for (int i = 1; i < args.Length; i++)
            {
                CommandRunner.Run((CommandRunner.Command)JsonConvert.DeserializeObject(args[i]));
            }
        }
    }
}
