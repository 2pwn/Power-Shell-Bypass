using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.IO;
using ClosedXML;
using System.Reflection;
using ClosedXML.Excel;
//using Microsoft.Management.Infrastructure;

namespace PawanPal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {



            #region Win32 Services

            ManagementScope scope = new ManagementScope("\\\\localhost\\root\\cimv2");
            scope.Connect();
            ObjectQuery query = new ObjectQuery("SELECT * FROM win32_service");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            //foreach (ManagementObject service in searcher.Get())
            //{
            //    // show the service
            //   label1.Text= service.ToString()+"</br>";
            //}

           // ManagementObjectSearcher s =
           //new ManagementObjectSearcher(
           //"localhost\\root\\CIMV2",
           //"select PSComputername, Name, InstallDate, DisplayName, PathName, State, StartMode from win32_systemdriver",
           //new EnumerationOptions(
           //null, System.TimeSpan.MaxValue,
           //1, true, false, true,
           //true, false, true, true));


           // foreach (ManagementObject service in s.Get())
           // {
           //     // show the service
           //     label1.Text = service.ToString() + "</br>";
           // }
           
            ManagementObjectCollection queryCollection = searcher.Get();
           // List<ManagementObject> managementList = queryCollection.Cast<ManagementObject>().ToList();
            //foreach (ManagementObject m in queryCollection)
            //{
            //    // Display the remote computer information
            //    Console.WriteLine("Computer Name     : {0}", m["csname"]);
            //    Console.WriteLine("Windows Directory : {0}", m["WindowsDirectory"]);
            //    Console.WriteLine("Operating System  : {0}", m["Caption"]);
            //    Console.WriteLine("Version           : {0}", m["Version"]);
            //    Console.WriteLine("Manufacturer      : {0}", m["Manufacturer"]);
            //    label1.Text = m["csname"].ToString() + "-" + m["WindowsDirectory"].ToString() + "-" + m["Caption"].ToString() + "<br/>" + m["Version"].ToString() + "<br/>" + m["Manufacturer"].ToString() + "</br>";
            //}


            List<Win32Service> objWin32Service = new  List<Win32Service>();
            List<win32_startupcommand> objwin32_startupcommand = new List<win32_startupcommand>();
            int i=1;
            foreach (ManagementObject m in queryCollection)
            {
                Win32Service objsinWin32Service = new Win32Service();
                // Display the remote computer information
                //Console.WriteLine("Computer Name     : {0}", m["csname"]);
                objsinWin32Service.Name = m["Name"] != null ? m["Name"].ToString() : null;
                objsinWin32Service.ID =i;
                //label1.Text = m["Name"] != null ? m["Name"].ToString() : null;
                objsinWin32Service.ProcessId = m["ProcessId"] != null ? m["ProcessId"].ToString() : null;
                objsinWin32Service.PathName = m["PathName"] != null ? m["PathName"].ToString() : null;
                objsinWin32Service.ServiceType = m["ServiceType"] != null ? m["ServiceType"].ToString() : null;
                objsinWin32Service.Started = m["Started"] != null ? m["Started"].ToString() : null;
                objsinWin32Service.State = m["State"] != null ? m["State"].ToString() : null;
                objsinWin32Service.Description = m["Description"] != null ? m["Description"].ToString() : null;
                objsinWin32Service.StartMode = m["StartMode"] != null ? m["StartMode"].ToString() : null;
                objsinWin32Service.StartName = m["StartName"] != null ? m["StartName"].ToString() : null;
                objsinWin32Service.PSComputerName = m["SystemName"] != null ? m["SystemName"].ToString() : null;


                i++;
                objWin32Service.Add(objsinWin32Service);
            }


            #endregion


            #region Win32 StartupCommand
            ObjectQuery query1 = new ObjectQuery("SELECT * FROM win32_startupcommand");
            ManagementObjectSearcher searcher1 = new ManagementObjectSearcher(scope, query1);
           queryCollection = searcher1.Get();
           
            foreach (ManagementObject m in queryCollection)
            {
                win32_startupcommand objsinwin32_startupcommand = new win32_startupcommand();
                // Display the remote computer information
                //Console.WriteLine("Computer Name     : {0}", m["csname"]);
                objsinwin32_startupcommand.Name = m["Name"] != null ? m["Name"].ToString() : null;
                //label1.Text = m["Name"] != null ? m["Name"].ToString() : null;
                objsinwin32_startupcommand.User = m["User"] != null ? m["User"].ToString() : null;
                objsinwin32_startupcommand.Location = m["Location"] != null ? m["Location"].ToString() : null;
                objsinwin32_startupcommand.Command = m["Command"] != null ? m["Command"].ToString() : null;
                //objsinwin32_startupcommand.PSComputerName = m["SystemName"] != null ? m["SystemName"].ToString() : null;



                objwin32_startupcommand.Add(objsinwin32_startupcommand);
            }
            dataGridView1.DataSource = objWin32Service;
            string folderPath = @"C:\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ToDataTable(objwin32_startupcommand), "win32_startupcommand");
                wb.SaveAs(folderPath + "objwin32_startupcommand.xlsx");
            };
            int a = 10;


            #endregion

            // string Namespace = @"root\cimv2";
       //   string OSQuery = "SELECT * FROM Win32_OperatingSystem";
       //CimSession mySession = CimSession.Create("Computer_B");
       // IEnumerable<CimInstance> queryInstance = mySession.QueryInstances(Namespace, "WQL", OSQuery);
            ////Get-WmiObject -Class win32_systemdriver -ComputerName $computers | select PSComputername, Name, InstallDate, DisplayName, PathName, State, StartMode | Export-CSV .\Drivers.csv -NoTypeInformation


            #region win32_networkloginprofile

            ObjectQuery qry_win32_networkloginprofile = new ObjectQuery("SELECT * FROM win32_networkloginprofile");
            ManagementObjectSearcher searcher_win32_networkloginprofile = new ManagementObjectSearcher(scope, qry_win32_networkloginprofile);
            queryCollection = searcher_win32_networkloginprofile.Get();
            List<win32_networkloginprofile> obj_win32_networkloginprofile_List = new List<win32_networkloginprofile>();

            foreach (ManagementObject m in queryCollection)
            {
                win32_networkloginprofile obj_win32_networkloginprofile = new win32_networkloginprofile();
                // Display the remote computer information
                //Console.WriteLine("Computer Name     : {0}", m["csname"]);
                obj_win32_networkloginprofile.Name = m["Name"] != null ? m["Name"].ToString() : null;
                //label1.Text = m["Name"] != null ? m["Name"].ToString() : null;
                obj_win32_networkloginprofile.PSComputerName = m["Workstations"] != null ? m["Workstations"].ToString() : null;
                obj_win32_networkloginprofile.LastLogon = m["LastLogon"] != null ? m["LastLogon"].ToString() : null;
                obj_win32_networkloginprofile.LastLogoff = m["LastLogoff"] != null ? m["LastLogoff"].ToString() : null;
                obj_win32_networkloginprofile.NumberOfLogons = m["NumberOfLogons"] != null ? m["NumberOfLogons"].ToString() : null;
                obj_win32_networkloginprofile.PasswordAge = m["PasswordAge"] != null ? m["PasswordAge"].ToString() : null;


                obj_win32_networkloginprofile_List.Add(obj_win32_networkloginprofile);
              
            }
            
            string folderPath1 = @"C:\Users\Pankaj Gupta\Desktop\Excel\";
            if (!Directory.Exists(folderPath1))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ToDataTable(obj_win32_networkloginprofile_List), "win32_networkloginprofile");
                wb.SaveAs(folderPath + "obj_win32_networkloginprofile.xlsx");
            };
           


            #endregion

            #region win32_logicaldisk

            ObjectQuery qry_win32_logicaldisk = new ObjectQuery("SELECT * FROM win32_logicaldisk");
            ManagementObjectSearcher searcher_win32_logicaldisk = new ManagementObjectSearcher(scope, qry_win32_logicaldisk);
            queryCollection = searcher_win32_logicaldisk.Get();
            List<win32_logicaldisk> obj_win32_logicaldisk = new List<win32_logicaldisk>();

            foreach (ManagementObject m in queryCollection)
            {
                win32_logicaldisk obj_win32_logicaldiskClss = new win32_logicaldisk();
                // Display the remote computer information

                obj_win32_logicaldiskClss.PSComputerName = m["SystemName"] != null ? m["SystemName"].ToString() : null;
                obj_win32_logicaldiskClss.DeviceID = m["DeviceID"] != null ? m["DeviceID"].ToString() : null;
                obj_win32_logicaldiskClss.Description = m["Description"] != null ? m["Description"].ToString() : null;
                obj_win32_logicaldiskClss.ProviderName = m["ProviderName"] != null ? m["ProviderName"].ToString() : null;


                obj_win32_logicaldisk.Add(obj_win32_logicaldiskClss);

            }

            string folderPath_win32_logicaldisk = @"C:\Users\Pankaj Gupta\Desktop\Excel\";
            if (!Directory.Exists(folderPath_win32_logicaldisk))
            {
                Directory.CreateDirectory(folderPath_win32_logicaldisk);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ToDataTable(obj_win32_logicaldisk), "obj_win32_logicaldisk");
                wb.SaveAs(folderPath_win32_logicaldisk + "obj_win32_logicaldisk.xlsx");
            };
           

            #endregion


            #region win32_process

            ObjectQuery qry_win32_process = new ObjectQuery("SELECT * FROM win32_process");
            ManagementObjectSearcher searcher_win32_process = new ManagementObjectSearcher(scope, qry_win32_process);
            queryCollection = searcher_win32_process.Get();
            List<win32_process> obj_win32_process = new List<win32_process>();

            foreach (ManagementObject m in queryCollection)
            {
                win32_process obj_win32_processClss = new win32_process();
                // Display the remote computer information

              //  obj_win32_processClss.PSComputerName = m["SystemName"] != null ? m["SystemName"].ToString() : null;
                obj_win32_processClss.Name = m["Name"] != null ? m["Name"].ToString() : null;
                obj_win32_processClss.Description = m["Description"] != null ? m["Description"].ToString() : null;
                obj_win32_processClss.ProcessId = m["ProcessId"] != null ? m["ProcessId"].ToString() : null;
                obj_win32_processClss.ParentProcessID = m["ParentProcessID"] != null ? m["ParentProcessID"].ToString() : null;
                obj_win32_processClss.Handle = m["Handle"] != null ? m["Handle"].ToString() : null;
                obj_win32_processClss.HandleCount = m["HandleCount"] != null ? m["HandleCount"].ToString() : null;
                obj_win32_processClss.ThreadCount = m["ThreadCount"] != null ? m["ThreadCount"].ToString() : null;
                obj_win32_processClss.CreationDate = m["CreationDate"] != null ? m["CreationDate"].ToString() : null;


                obj_win32_process.Add(obj_win32_processClss);

            }

            string folderPath_win32_process = @"C:\Users\Pankaj Gupta\Desktop\Excel\";
            if (!Directory.Exists(folderPath_win32_process))
            {
                Directory.CreateDirectory(folderPath_win32_process);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ToDataTable(obj_win32_process), "obj_win32_process");
                wb.SaveAs(folderPath_win32_process + "obj_win32_process.xlsx");
            };
           

            #endregion


            #region win32_systemdriver

            ObjectQuery qry_win32_systemdriver = new ObjectQuery("SELECT * FROM win32_systemdriver");
            ManagementObjectSearcher searcher_win32_systemdriver = new ManagementObjectSearcher(scope, qry_win32_systemdriver);
            queryCollection = searcher_win32_systemdriver.Get();
            List<win32_systemdriver> obj_win32_systemdriver = new List<win32_systemdriver>();

            foreach (ManagementObject m in queryCollection)
            {
                win32_systemdriver obj_win32_systemdriverClss = new win32_systemdriver();
                // Display the remote computer information

                obj_win32_systemdriverClss.PSComputerName = m["SystemName"] != null ? m["SystemName"].ToString() : null;
                obj_win32_systemdriverClss.Name = m["Name"] != null ? m["Name"].ToString() : null;
                obj_win32_systemdriverClss.InstallDate = m["InstallDate"] != null ? m["InstallDate"].ToString() : null;
                obj_win32_systemdriverClss.DisplayName = m["DisplayName"] != null ? m["DisplayName"].ToString() : null;
                obj_win32_systemdriverClss.PathName = m["PathName"] != null ? m["PathName"].ToString() : null;
                obj_win32_systemdriverClss.State = m["State"] != null ? m["State"].ToString() : null;
                obj_win32_systemdriverClss.StartMode = m["StartMode"] != null ? m["StartMode"].ToString() : null;


                obj_win32_systemdriver.Add(obj_win32_systemdriverClss);

            }

            string folderPath_win32_systemdriver = @"C:\Users\Pankaj Gupta\Desktop\Excel\";
            if (!Directory.Exists(folderPath_win32_systemdriver))
            {
                Directory.CreateDirectory(folderPath_win32_systemdriver);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ToDataTable(obj_win32_systemdriver), "obj_win32_systemdriver");
                wb.SaveAs(folderPath_win32_systemdriver + "obj_win32_systemdriver.xlsx");
            };
           

            #endregion

            #region win32_environment

            ObjectQuery qry_win32_environment = new ObjectQuery("SELECT * FROM win32_environment");
            ManagementObjectSearcher searcher_win32_environment = new ManagementObjectSearcher(scope, qry_win32_environment);
            queryCollection = searcher_win32_environment.Get();
            List<win32_environment> obj_win32_environment = new List<win32_environment>();

            foreach (ManagementObject m in queryCollection)
            {
                win32_environment obj_win32_environmentClss = new win32_environment();
                // Display the remote computer information

                obj_win32_environmentClss.PSComputerName = m["Description"] != null ? m["Description"].ToString() : null;
                obj_win32_environmentClss.Name = m["Name"] != null ? m["Name"].ToString() : null;
                obj_win32_environmentClss.UserName = m["UserName"] != null ? m["UserName"].ToString() : null;
                obj_win32_environmentClss.VariableValue = m["VariableValue"] != null ? m["VariableValue"].ToString() : null;


                obj_win32_environment.Add(obj_win32_environmentClss);

            }

            string folderPath_win32_environment = @"C:\Users\Pankaj Gupta\Desktop\Excel\";
            if (!Directory.Exists(folderPath_win32_environment))
            {
                Directory.CreateDirectory(folderPath_win32_environment);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ToDataTable(obj_win32_environment), "obj_win32_environment");
                wb.SaveAs(folderPath_win32_environment + "obj_win32_environment.xlsx");
            };
           

            #endregion


            #region win32_networkadapterconfiguration

            ObjectQuery qry_win32_networkadapterconfiguration = new ObjectQuery("SELECT * FROM win32_networkadapterconfiguration");
            ManagementObjectSearcher searcher_win32_networkadapterconfiguration = new ManagementObjectSearcher(scope, qry_win32_networkadapterconfiguration);
            queryCollection = searcher_win32_networkadapterconfiguration.Get();
            List<win32_networkadapterconfiguration> obj_win32_networkadapterconfiguration = new List<win32_networkadapterconfiguration>();

            foreach (ManagementObject m in queryCollection)
            {
                win32_networkadapterconfiguration obj_win32_networkadapterconfigurationClss = new win32_networkadapterconfiguration();
                // Display the remote computer information

               // obj_win32_networkadapterconfigurationClss.PSComputerName = m["SystemName"] != null ? m["SystemName"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.IPAddress = m["IPAddress"] != null ? m["IPAddress"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.IPSubnet = m["IPSubnet"] != null ? m["IPSubnet"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.DefaultIPGateway = m["DefaultIPGateway"] != null ? m["DefaultIPGateway"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.DHCPServer = m["DHCPServer"] != null ? m["DHCPServer"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.DNSHostname = m["DNSHostname"] != null ? m["DNSHostname"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.DNSserversearchorder = m["DNSserversearchorder"] != null ? m["DNSserversearchorder"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.MACAddress = m["MACAddress"] != null ? m["MACAddress"].ToString() : null;
                obj_win32_networkadapterconfigurationClss.description = m["description"] != null ? m["description"].ToString() : null;



                obj_win32_networkadapterconfiguration.Add(obj_win32_networkadapterconfigurationClss);

            }

            string folderPath_win32_networkadapterconfiguration = @"C:\Users\Pankaj Gupta\Desktop\Excel\";
            if (!Directory.Exists(folderPath_win32_networkadapterconfiguration))
            {
                Directory.CreateDirectory(folderPath_win32_networkadapterconfiguration);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(ToDataTable(obj_win32_networkadapterconfiguration), "obj_win32_net_adapterconf");
                wb.SaveAs(folderPath_win32_networkadapterconfiguration + "obj_win32_net_adapterconf.xlsx");
            };
           

            #endregion

        }

        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties by using reflection   
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {

                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
