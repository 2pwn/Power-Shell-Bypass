using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PawanPal
{
   public class Win32Service
    {
       public int ID { get; set; }
       public string Name {get;set;}
           public string PathName {get;set;}
               public string Description {get;set;}
                   public string ProcessId {get;set;}
                       public string ServiceType {get;set;}
                            public string Started {get;set;}
                                   public string StartMode {get;set;}
                                       public string StartName {get;set;}
                                       public string State { get; set; }
                                       public string PSComputerName { get; set; }
                           
    }
}
