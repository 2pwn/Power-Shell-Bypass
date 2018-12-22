using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PawanPal
{
  public  class win32_networkloginprofile
    {
        public int ID { get; set; }
        public string Name { get; set; }

        public string LastLogon  { get; set; }
        public string LastLogoff     { get; set; }
        public string NumberOfLogons { get; set; }
        public string PasswordAge { get; set; }
      
      //public string DHCPServer { get; set; }
      //  public string DefaultIPGateway { get; set; }
      //  public string IPSubnet { get; set; }
      //  public string IpAddress { get; set; }
        public string PSComputerName { get; set; }
    }
}
