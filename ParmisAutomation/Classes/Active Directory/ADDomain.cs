using System;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices.ActiveDirectory;

namespace ParmisAutomation.Classes
{
    public class ADDomain : ParmisAutomation
    {

        //*************************  Class Attributes   *************************  ویژگی های کلاس

        private string _Name;
        private DomainControllerCollection _All_DC;
        private int _NumberOfDC;
        private const int _InactiveDay = 5;

        //*************************  Class Properties   *************************  خصوصیات کلاس
        public string Name
        {
            get
            {
                return _Name;
            }
        }
        public DomainControllerCollection All_DC
        {
            get
            {
                return _All_DC;
            }
        }
        public int NumberOfDC
        {
            get
            {
                return _NumberOfDC;
            }
        }

        //************************* Constructure *************************    متد سازنده کلاس

        public ADDomain()
        {
            Domain CurrentDomain = Domain.GetCurrentDomain();
            _Name = CurrentDomain.Name;            
            _All_DC = CurrentDomain.FindAllDomainControllers();
            _NumberOfDC = _All_DC.Count;
        }

        //************************* DisableInactiveUsers *************************   غیر فعال سازی کاربران قدیمی

        public Collection<PSObject> DisableInactiveUsers(string DC,string SearchLocation,int InactiveDay = _InactiveDay)
        {
            PShell Command = new PShell();
            Command.PSExec("Invoke-Command -ComputerName " + DC + " -ScriptBlock {\n" +
                        "$InactiveDay =" + InactiveDay.ToString() + "\n" +
                        "$SearchLOC = \"" + SearchLocation + "\" \n " +
                        "$BaseDay = (Get-Date).AddDays(-$InactiveDay)\n" +
                        "$AllUsers = Get-ADUser -SearchBase \"$SearchLOC\" -Filter { LastLogonTimeStamp -lt $BaseDay } \n" +
                        "$AllUsers | Disable-ADAccount \n" +
                        "$AllUsers \n" +
                        "}\n");
            return Command.Output;
        }     

        public Collection<PSObject> DisableInactiveUsers(string SearchLocation,int InactiveDay = _InactiveDay)
        {
            PShell Command = new PShell();
            foreach (DomainController DC in _All_DC)
            {                               
                if (PingTest(DC.Name) == "Success")
                {
                    Command.PSExec( "Invoke-Command -ComputerName " + DC.Name + " -ScriptBlock {\n" +
                                "$InactiveDay =" + InactiveDay.ToString() + "\n" +
                                "$SearchLOC = \"" + SearchLocation + "\" \n " +
                                "$BaseDay = (Get-Date).AddDays(-$InactiveDay)\n" +
                                "$AllUsers = Get-ADUser -SearchBase \"$SearchLOC\" -Filter { LastLogonTimeStamp -lt $BaseDay } \n" +
                                "$AllUsers | Disable-ADAccount \n" +
                                "$AllUsers \n" +
                                "}\n");                                            
                    if (Command.Status == "Success" && Command.Output.Count > 0)
                        break;
                }                         
            }            
            return Command.Output;
        }

        //************************* DisableInactiveComputers *************************   غیر فعال سازی کامپیوترهای قدیمی

        public Collection<PSObject> DisableInactiveComputers(string DC, string SearchLocation, int InactiveDay = _InactiveDay)
        {
            PShell Command = new PShell();
            Command.PSExec("Invoke-Command -ComputerName " + DC + " -ScriptBlock {\n" +
                                    "$InactiveDay =" + InactiveDay.ToString() + "\n" +
                                    "$SearchLOC = \"" + SearchLocation + "\" \n " +
                                    "$BaseDay = (Get-Date).AddDays(-$InactiveDay)\n" +
                                    "$AllComputer = Get-ADComputer -SearchBase \"$SearchLOC\" -Filter { LastLogonTimeStamp -lt $BaseDay } \n" +
                                    "$AllComputer | Disable-ADAccount \n" +
                                    "$AllComputer \n" +
                                    "}\n");
            return Command.Output;
        }

        public Collection<PSObject> DisableInactiveComputers(string SearchLocation, int InactiveDay = _InactiveDay)
        {
            PShell Command = new PShell();
            foreach (DomainController DC in _All_DC)
            {
                if (PingTest(DC.Name) == "Success")
                {
                    Command.PSExec("Invoke-Command -ComputerName " + DC.Name + " -ScriptBlock {\n" +
                                "$InactiveDay =" + InactiveDay.ToString() + "\n" +
                                "$SearchLOC = \"" + SearchLocation + "\" \n " +
                                "$BaseDay = (Get-Date).AddDays(-$InactiveDay)\n" +
                                "$AllComputers = Get-ADComputer -SearchBase \"$SearchLOC\" -Filter { LastLogonTimeStamp -lt $BaseDay } \n" +
                                "$AllComputers | Disable-ADAccount \n" +
                                "$AllComputers \n" +
                                "}\n");
                    if (Command.Status == "Success" && Command.Output.Count > 0)
                        break;
                }
            }
            return Command.Output;
        }

        //************************* RemoveDisabledUsers *************************   حذف کاربران غیر فعال

        public Collection<PSObject> RemoveDisabledUsers(string DC, string SearchLocation)
        {
            PShell Command = new PShell();
            Command.PSExec("Invoke-Command -ComputerName " + DC + " -ScriptBlock {\n" +
                                    "$AllUsers = Search-ADAccount -SearchBase \"" + SearchLocation + "\" -AccountDisabled | Where-Object {$_.ObjectClass -eq \"user\" } \n" +
                                    "$AllUsers | Remove-ADUser -confirm:$false \n" +
                                    "$AllUsers \n" +
                                    "}\n");
            return Command.Output;
        }

        public Collection<PSObject> RemoveDisabledUsers(string SearchLocation)
        {
            PShell Command = new PShell();
            foreach (DomainController DC in _All_DC)
            {
                if (PingTest(DC.Name) == "Success")
                {
                    Command.PSExec("Invoke-Command -ComputerName " + DC + " -ScriptBlock {\n" +
                                            "$AllUsers = Search-ADAccount -SearchBase \"" + SearchLocation + "\" -AccountDisabled | Where-Object {$_.ObjectClass -eq \"user\" } \n" +
                                            "$AllUsers | Remove-ADUser -confirm:$false \n" +
                                            "$AllUsers \n" +
                                            "}\n");
                    if (Command.Status == "Success" && Command.Output.Count > 0)
                        break;
                }
            }
            return Command.Output;
        }

        //************************* RemoveDisabledComputers *************************   حذف کامپیوترهای غیر فعال

        public Collection<PSObject> RemoveDisabledComputers(string DC, string SearchLocation)
        {
            PShell Command = new PShell();
            Command.PSExec("Invoke-Command -ComputerName " + DC + " -ScriptBlock {\n" +
                                    "$AllComputers = Search-ADAccount -SearchBase \"" + SearchLocation + "\" -AccountDisabled | Where-Object {$_.ObjectClass -eq \"computer\" } \n" +
                                    "$AllComputers | Remove-ADComputer -confirm:$false \n" +
                                    "$AllComputers \n" +
                                    "}\n");
            return Command.Output;
        }

        public Collection<PSObject> RemoveDisabledComputers(string SearchLocation)
        {
            PShell Command = new PShell();
            foreach (DomainController DC in _All_DC)
            {
                if (PingTest(DC.Name) == "Success")
                {
                    Command.PSExec("Invoke-Command -ComputerName " + DC + " -ScriptBlock {\n" +
                                            "$AllComputers = Search-ADAccount -SearchBase \"" + SearchLocation + "\" -AccountDisabled | Where-Object {$_.ObjectClass -eq \"computer\" } \n" +
                                            "$AllComputers | Remove-ADComputer -confirm:$false \n" +
                                            "$AllComputers \n" +
                                            "}\n");
                    if (Command.Status == "Success" && Command.Output.Count > 0)
                        break;
                }
            }
            return Command.Output;
        }

    }
}