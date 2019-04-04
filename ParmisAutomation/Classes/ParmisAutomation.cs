
using System;
using System.DirectoryServices.ActiveDirectory;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.NetworkInformation;


namespace ParmisAutomation
{
    public class ParmisAutomation
    {
        private Forest _CurrentForest;
        private Domain _CurrentDomain;

        //****************** خصوصیات کلاس *********************

        public Forest CurrentForest
        {
            get
            {
                return _CurrentForest;
            }
        }

        public Forest CurrentDomain
        {
            get
            {
                return _CurrentForest;
            }
        }

        //********************* متد سازنده کلاس ******************************

        public ParmisAutomation()
        {
            try
            {
                _CurrentForest = Forest.GetCurrentForest();
                _CurrentDomain = Domain.GetCurrentDomain();
            }
            catch
            {
                _CurrentForest = null;
                _CurrentDomain = null;
            }
        }

        //********************* توابع کلاس ******************************

        public string PingTest(string Host)
        {
            Ping P =new Ping();
            PingReply R;
            try
            {
                R = P.Send(Host);
                return R.Status.ToString();
            }
            catch (Exception)
            {
                return ("TimedOut");
            }
        }
        
        //********************* توابع کلاس ****************************** 
    }
}
