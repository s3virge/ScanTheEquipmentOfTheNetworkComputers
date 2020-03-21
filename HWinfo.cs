using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Management;

namespace ScanTheEquipmentOfTheNetworkComputers {
    class HWinfo
    {
        private ManagementScope _scope;
        private string _hostName = "";      

        /// <summary>
        /// constructor that initializes _hostname field
        /// </summary>
        /// <param name="computerToConnect"></param>
        public HWinfo(string computerToConnect = "")
        {           
            _hostName = computerToConnect;            
        }

        /// <summary>
        /// Use _hostName field for connect
        /// </summary>
        //public void Connect()
        //{
        //    ConnectionOptions options = new ConnectionOptions();
        //    options.Impersonation = ImpersonationLevel.Impersonate;

        //    _scope = new ManagementScope($@"\\{_hostName}\root\cimv2", options);

        //    _scope.Connect();
        //}

        /// <summary>
        /// Use certain computer name for connect. If okay _scope field will be 
        /// initialised
        /// </summary>
        /// <param name="computerNameToConnect"></param>
        public void Connect(string computerNameToConnect = "")
        {
            try
            {
                if (computerNameToConnect == "")
                {
                    computerNameToConnect = _hostName;
                }

                ConnectionOptions options = new ConnectionOptions();
                options.Impersonation = ImpersonationLevel.Impersonate;

                _scope = new ManagementScope($@"\\{computerNameToConnect}\root\cimv2", options);

                _scope.Connect();
            }
            catch
            {
                Debug.WriteLine($"Облом с подключением к компу {computerNameToConnect}");
            }
            
        }      

        private ManagementObjectCollection GetManagementObjectCollection(string query)
        {
            ManagementObjectCollection queryCollection = null;
            try
            {
                ObjectQuery objQuery = new ObjectQuery(query);
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(_scope, objQuery);
                queryCollection = searcher.Get();             
            }
            catch (Exception)
            {                
                throw new Exception("Облом при получении ManagementObjectCollection");
            }

            return queryCollection;
        }

        DateTime ConvertToDateTime(string str)
        {
            //2019 12 03 10 18 27 .500000+120
            int year = Convert.ToInt32(str.Substring(0, 4));
            int month = Convert.ToInt32(str.Substring(4, 2));
            int day = Convert.ToInt32(str.Substring(6, 2));

            return new DateTime(year, month, day);
        }

        public string GetComputerName()
        {
            string result = "Unknown Computer Name";

            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_OperatingSystem");

            foreach (ManagementObject m in queryCollection)
            {
                result = String.Format(Output.format, m["csname"]);
                
            }

            return result;
        }

        public string GetOperatingSystem()
        {
            string result = "Unknown Operating System";
            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_OperatingSystem");

            foreach (ManagementObject m in queryCollection)
            {
                result = String.Format(Output.format, m["Caption"]);
                
            }

            return result;
        }

        public string GetMotherBoard()
        {
            string result = "Unknown MotherBoard";
            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_BaseBoard");

            foreach (ManagementObject m in queryCollection)
            {
                result = String.Format(Output.format, m["Manufacturer"]);
                result += " " + String.Format(Output.format, m["Product"]);
                result = result.Replace(",", "");
                //Debug.WriteLine(result);
            }

            return result;
        }

        public string GetProcessor()
        {
            string result = "Unknown Processor";

            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_Processor");

            foreach (ManagementObject m in queryCollection)
            {
                result = String.Format(Output.format, m["Name"]);
                //result += String.Format(Output.format, "\t", m["Caption"]);
                //result;
                
            }

            return result;
        }

        public string GetMemory()
        {
            string result = "Unknown Memory";
            string memoryOutputFormat = Output.format + " GB";
            List<string> res = new List<string>(4) {result, result, result, result };
            
            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_PhysicalMemory");

            int index = 0;
            foreach (ManagementObject m in queryCollection)
            {
                result = String.Format(Output.format, m["PartNumber"]);
                Int64 memoryCapasity = Convert.ToInt64(m["Capacity"]) / 1024 / 1024 / 1024;
                result += " " + String.Format(memoryOutputFormat, memoryCapasity);

                res[index] = result;
                index++;
            }
            
            result = string.Join(",", res.ToArray());

            return result ;
        }

        public string GetDiskDrive()
        {
            string result = "Unknown Disk Drive";
            List<string> res = new List<string>(4) { result, result, result, result };

            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_DiskDrive");

            //string driveOutputFormat = Output.format + " GB";
            int index = 0;
            foreach (ManagementObject m in queryCollection)
            {
                result = String.Format(Output.format, m["Model"]);
                //Int64 driveSize = Convert.ToInt64(m["Size"]) / 1024 / 1024 / 1024;
                //result += String.Format(driveOutputFormat, "\t", driveSize);
                //result;
                
                res[index] = result;
                index++;
            }

            return result;
        }

        public string GetVideoController()
        {
            string result = "Unknown VideoController";

            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_VideoController");

            foreach (ManagementObject m in queryCollection)
            {
                //result += String.Format(outputFormat, "Caption", m["Caption"]);
                //result = String.Format(Output.format, "Video", m["Name"]);                
                result = String.Format(Output.format, m["VideoProcessor"]);
                //result;
                //result += String.Format(Output.format, "\t", m["VideoModeDescription"]);
                //result;
                
            }

            return result;
        }

        public string GetNetworkAdapter()
        {
            string result = "Unknown NetworkAdapter";

            ManagementObjectCollection queryCollection = GetManagementObjectCollection("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'");

            foreach (ManagementObject m in queryCollection)
            {
                //result += String.Format(outputFormat, "Caption", m["Caption"]);
                result = String.Format(Output.format, m["Description"]);

                //String[] ip = (String[])m["IPAddress"];

                //result += String.Format(Output.format, "\t", ip[0]);
                //result;
                
            }

            return result;
        }

        /// <summary>
        /// get Windows releaseID (1803, 1903, 1909 ...) from remote pc
        /// </summary>
        /// <returns></returns>
        private string GetOperatingSystemReleaseID()
        {
            string version = "";

            try
            {
                using (RegistryKey remoteHklm = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, _hostName))
                {
                    using (RegistryKey serviceKey = remoteHklm.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", false))
                    {
                        if (serviceKey != null)
                        {
                            version = serviceKey.GetValue("ReleaseId").ToString();
                        }
                        else
                        {
                            version = "error on get version from registry";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }

            return version;
        }

        public string Get()
        {
            string listOfHardware = "";

            try
            {
                listOfHardware += GetComputerName() + ",";               
                listOfHardware += GetOperatingSystem() + ",";
                listOfHardware += GetMotherBoard() + ",";
                listOfHardware += GetProcessor() + ",";
                listOfHardware += GetMemory() + ",";
                listOfHardware += GetDiskDrive() + ",";
                listOfHardware += GetVideoController() + ",";
                listOfHardware += GetNetworkAdapter();                
            }
            catch(Exception ex)
            {
                Debug.WriteLine($"Get() - Получить инфу не удалось. {ex.Message}");
            }

            return listOfHardware;
        }
    }
}
