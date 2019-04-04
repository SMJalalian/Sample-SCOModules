using System;
using Microsoft.SystemCenter.Orchestrator.Integration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParmisAutomation.Classes.Active_Directory.Orchestrator
{
    [Activity("Disable Inactive Computers")]
    class DisableInactiveComputers : ADDomain
    {
        private string _Orc_DC;
        private string _Orc_SearchLoc;
        private int _Orc_InactiveDays;
        private bool _Orc_Result;

        // ********************************** Activity Input ***********************************

        [ActivityInput("Name of DC")]
        public string Orc_DC { set { _Orc_DC = value; } get { return _Orc_DC; } }

        [ActivityInput("Distinguished Name of OU")]
        public string Orc_SearchLoc { set { _Orc_SearchLoc = value; } get { return _Orc_SearchLoc; } }

        [ActivityInput("Number of Inactive Days")]
        public int Orc_InactiveDays { set { _Orc_InactiveDays = value; } get { return _Orc_InactiveDays; } }

        // ********************************** Activity Output ***********************************

        [ActivityOutput("Execute Result", Description = "if Result = Sucess, Means that Script Execute Successfully")]
        public bool Orc_Result { set { _Orc_Result = value; } get { return _Orc_Result; } }

        // ********************************** Activity Methode ***********************************

        [ActivityMethod]
        public void Run()
        {
            try
            {
                if (Orc_DC != null)
                    DisableInactiveComputers(Orc_DC, Orc_SearchLoc, Orc_InactiveDays);
                else
                    DisableInactiveComputers(Orc_SearchLoc, Orc_InactiveDays);
                Orc_Result = true;
            }
            catch (Exception)
            {
                Orc_Result = false;
            }

        }
    }
}
