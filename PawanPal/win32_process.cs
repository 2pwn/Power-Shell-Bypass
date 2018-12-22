using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PawanPal
{
   public class win32_process
    {
        public int ID { get; set; }
        public string PSComputerName { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string ProcessId { get; set; }

        public string ParentProcessID { get; set; }

        public string Handle { get; set; }
        public string HandleCount { get; set; }
        public string ThreadCount { get; set; }
        public string CreationDate { get; set; }
      
    }
}
