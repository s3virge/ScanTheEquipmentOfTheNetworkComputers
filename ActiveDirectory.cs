using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Net;
using System.Text.RegularExpressions;

namespace ScanTheEquipmentOfTheNetworkComputers {
    class ActiveDirectory
    {
        private string _domainOU;

        /// <summary>
        /// set domain OU for saerch
        /// </summary>
        /// <param name="domainOU">"DC=intetics,DC=com,DC=ua"</param>
        public ActiveDirectory()
        {
            try
            {
                _domainOU = Domain.GetComputerDomain().ToString();
            }
            catch (Exception)
            {
                throw new Exception("This computer do not joined to domain");
            }
        }

        /// <summary>
        /// receive from AD list of computers with given name
        /// </summary>
        /// <param name="comuter Name"></param>
        /// <returns>array list whith computers names</returns>
        public ArrayList GetListOfComputers(string comuterName = null)
        {
            DirectoryEntry entry = new DirectoryEntry($"LDAP://{_domainOU}");
            entry.RefreshCache();
            DirectorySearcher mySearcher = new DirectorySearcher(entry);

            ArrayList compList = new ArrayList();

            //search parameters
            //mySearcher.Filter = $"( &(objectClass=computer)(Name=*{comuterName}*)(cn=*{comuterName}*))";

            if (comuterName == null || comuterName == "")
            {
                mySearcher.Filter = $"( &(objectClass=computer))";
            }
            else
            {
                mySearcher.Filter = $"( &(objectClass=computer)(cn=*{comuterName}*))";
            }

            mySearcher.SizeLimit = int.MaxValue;
            mySearcher.PageSize = int.MaxValue;

            foreach (SearchResult resEnt in mySearcher.FindAll())
            {
                //"CN=SGSVG007DC"
                DirectoryEntry directoryEntry = new DirectoryEntry();
                directoryEntry = resEnt.GetDirectoryEntry();

                //string sAMAccountName = directoryEntry.Properties["sAMAccountName"].Value.ToString();
                if (IsActive(directoryEntry))
                {
                    string ComputerName = directoryEntry.Name;

                    if (ComputerName.StartsWith("CN="))
                    {
                        ComputerName = ComputerName.Remove(0, "CN=".Length);
                    }

                    compList.Add(ComputerName);
                }
            }

            mySearcher.Dispose();
            entry.Dispose();

            return compList;
        }

        /// <summary>
        /// check if derectory entry is disabled
        /// </summary>
        /// <param name="dirEntry"></param>
        /// <returns></returns>
        private bool IsActive(DirectoryEntry dirEntry)
        {
            if (dirEntry.NativeGuid == null)
                return false;

            int flags = (int)dirEntry.Properties["userAccountControl"].Value;

            return !Convert.ToBoolean(flags & 0x0002);
        }

        /// <summary>
        /// retrieve form Active Directory information about selected computer
        /// </summary>      
        public string GetComputerInfo(string computer)
        {
            string adComputerInfo = "";

            DirectoryEntry entry = new DirectoryEntry("LDAP://" + _domainOU);
            //Инициализирует новый экземпляр класса DirectorySearcher с указанным корнем и фильтром поиска, а также с указанием извлекаемых свойств.
            string searchfilter = $"(&(objectClass=computer)(Name={computer}))";
            //string searchfilter = "(&(objectClass=computer))";
            DirectorySearcher mySearcher = new DirectorySearcher(entry, searchfilter
                , new string[] { "name", "whenCreated", "whenChanged", "canonicalName", "OperatingSystem", "CanonicalName" }
                //,new string[] {"*"}
                );

            //mySearcher.PropertiesToLoad.Add("LastLogonDate");

            SearchResult resEnt = mySearcher.FindOne();

            if (resEnt != null)
            {
                // Note: Properties can contain multiple values.
                string name = "";
                string whenCreated = "";
                string whenChanged = "";
                string OperatingSystem = "";
                string CanonicalName = "";

                name = (string)resEnt.Properties["name"][0];
                whenCreated = (string)resEnt.Properties["whenCreated"][0].ToString();
                whenChanged = (string)resEnt.Properties["whenChanged"][0].ToString();
                CanonicalName = (string)resEnt.Properties["CanonicalName"][0].ToString();

                if (resEnt.Properties["OperatingSystem"].Count > 0)
                {
                    OperatingSystem = (string)resEnt.Properties["OperatingSystem"][0];
                }

                //if (resEnt.Properties["distinguishedname"].Count > 0) {
                //    distinguishedname = (string)resEnt.Properties["distinguishedname"][0];
                //    distinguishedname = distinguishedname.Remove(0, distinguishedname.IndexOf(",") + 1);
                //    //distinguishedname = distinguishedname.Replace("OU=", "").Replace("DC=", "").Replace(",", " ");
                //}

                //int flags = (int)resEnt.Properties["userAccountControl"][0];
                //if (Convert.ToBoolean(flags & 0x0002)) {
                //    enabled = "disabled";
                //}
                               
                //adComputerInfo += string.Format(Output.format, "OperatingSystem", $"{OperatingSystem}\n");
                //adComputerInfo += string.Format(OutputFormat.format, "Op Unit", $"{distinguishedname}\n");
                adComputerInfo += string.Format(Output.format, "CanonicalName", $"{CanonicalName}\n");
                //adComputerInfo += string.Format(Output.format, "whenCreated", $"{whenCreated}\n");
                //adComputerInfo += string.Format(Output.format, "whenChanged", $"{whenChanged}\n");
                //adComputerInfo += string.Format(Output.format, "IPv4Address", $"{GetDNSIPv4Address(computer)}\n");                
            }

            return adComputerInfo;
        }

        private string GetDNSIPv4Address(string hostName)
        {
            string dnsIPv4Address = "";
            IPHostEntry hostEntry = null;

            try
            {
                hostEntry = Dns.GetHostEntry(hostName);
                //you might get more than one ip for a hostname since 
                //DNS supports more than one record
                if (hostEntry.AddressList.Length > 0)
                {
                    dnsIPv4Address += hostEntry.AddressList[0];
                }
            }
            catch
            {
                Console.WriteLine($"An error has occurred in GetDNSIPv4Address()");
            }

            return dnsIPv4Address;
        }
    }
}
