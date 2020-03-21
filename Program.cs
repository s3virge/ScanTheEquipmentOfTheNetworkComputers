using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScanTheEquipmentOfTheNetworkComputers
{
    class Program
    {
        private static List<string> _allCompInfo = new List<string>();
        private static int count = 1;
        public static object lockObj = new object();

        static string GetInfo(string hostNameToConnect)
        {
            string result = "";

            try
            {
                HWinfo remoteHost = new HWinfo(hostNameToConnect);
                remoteHost.Connect();
                result = remoteHost.Get();
                
                lock (lockObj)
                {
                    Console.SetCursorPosition(0, 5);
                    Console.Write($"Complited" +
                        $": {count++}. Please wait.");
                }
            }
            catch (Exception ex)
            {
                result = $"Can not connect to {hostNameToConnect}: {ex.Message}";
            }

            return result;
        }

        // определение асинхронного метода
        static async void GetHWInfoAsync(string host)
        {
            string output = await Task.Run(() => GetInfo(host));

            //Console.Write($"{output}{ Output.lineSeparator}");           
            //lock (lockObj)
            { 
                Console.SetCursorPosition(0, 4);
                Console.Write($"Hosts complited: {count++}. Please wait.");
                _allCompInfo.Add(output);
            }            
        }

        static void Main(string[] args)
        {
            ActiveDirectory ad = new ActiveDirectory();
            ArrayList computers = ad.GetListOfComputers();

            Console.Write($"{Output.lineSeparator}");
            Console.WriteLine($"Number of hosts: {computers.Count}");
            Console.WriteLine($"{Output.lineSeparator}");            
            Console.WriteLine("Retrieving the hardware information from remote hosts.");
           
            List<Task> tasks = new List<Task>();
            
            int c = 0;
            foreach (string computer in computers)
            {
                //GetHWInfoAsync(computer as string);
                //tasks.Add(Task.Factory.StartNew(() => GetHWInfoAsync(computer)));
                tasks.Add(Task.Factory.StartNew(
                    () => {
                        _allCompInfo.Add(GetInfo(computer));                                            
                    }));
                //if (c >= 10) break;
                c++;
            }

            Task.WaitAll(tasks.ToArray());
            
            Console.WriteLine("\nAll done.");
                     
            
            WriteToFile();
            Console.WriteLine("Press eny key for exit.");
            Console.Read();
        }

        static void WriteToFile()
        {
            string file = @"HostsHWInfo.csv";

            try
            {
                using (StreamWriter sw = new StreamWriter(file, false, System.Text.Encoding.Default))
                {
                    foreach (var info in _allCompInfo)
                    {
                        sw.WriteLine($"{info}");
                    }
                }               
                
                Console.WriteLine($"Hardware information was saved to file {file}.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
