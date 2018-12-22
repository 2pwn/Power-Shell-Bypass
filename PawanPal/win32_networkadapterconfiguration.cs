using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PawanPal
{
   public class win32_networkadapterconfiguration
    {
        public int ID { get; set; }
        public string PSComputerName { get; set; }
        public string IPAddress { get; set; }
        public string IPSubnet { get; set; }
        public string DefaultIPGateway { get; set; }

        public string DHCPServer { get; set; }
        public string DNSHostname { get; set; }
        public string DNSserversearchorder { get; set; }
        public string MACAddress { get; set; }
        public string description { get; set; }

    }

}
