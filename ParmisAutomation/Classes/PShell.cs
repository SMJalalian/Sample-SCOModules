using System;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParmisAutomation.Classes
{
    public class PShell : ParmisAutomation
    {

        //*************************  Class Attributes   *************************  ویژگی های کلاس

        private Collection<PSObject> _Output;
        private Collection<ErrorRecord> _ErrorRecords;
        private Collection<WarningRecord> _WarningRecords;
        private Collection<InformationRecord> _InformationRecords;
        private PowerShell PS_Script = PowerShell.Create();
        private string _Status;

        //*************************  Class Properties   *************************  خصوصیات کلاس

        public Collection<PSObject> Output
        {
            get
            {
                return _Output;
            }
        }
        public Collection<ErrorRecord> ErrorRecords
        {
            get
            {
                return _ErrorRecords;
            }
        }
        public Collection<WarningRecord> WarningRecords
        {
            get
            {
                return _WarningRecords;
            }
        }
        public Collection<InformationRecord> InformationRecords
        {
            get
            {
                return _InformationRecords;
            }
        }
        public string Status
        {
            get
            {
                return _Status;
            }
        }

        //************************* Constructure *************************    متد سازنده کلاس

        public PShell()
        {
            _Output = null;
            _ErrorRecords = null;
            _WarningRecords = null;
            _InformationRecords = null;
            _Status = null;
        }

        public PShell(string _Script)
        {
            PS_Script.AddScript(_Script);
            try
            {
                _Output = PS_Script.Invoke();
                _ErrorRecords = PS_Script.Streams.Error.ReadAll();
                _WarningRecords = PS_Script.Streams.Warning.ReadAll();
                _InformationRecords = PS_Script.Streams.Information.ReadAll();
                if (_ErrorRecords != null && _ErrorRecords.Count >= 1)
                    _Status = "Error";
                else if (_WarningRecords != null && _WarningRecords.Count >= 1)
                    _Status = "Warning";
                else if (_InformationRecords != null && _InformationRecords.Count >= 1)
                    _Status = "Information";
                else
                    _Status = "Success";
            }
            catch (Exception)
            {
                _ErrorRecords = PS_Script.Streams.Error.ReadAll();
                _WarningRecords = PS_Script.Streams.Warning.ReadAll();
                _InformationRecords = PS_Script.Streams.Information.ReadAll();
                _Status = "Error";
                _Output = null;
            }
        }

        //************************* Methods *************************    متد های کلاس
           
        public void PSExec ( string _Script )
        {
            PS_Script.AddScript(_Script);
            try
            {
                _Output = PS_Script.Invoke();
                _ErrorRecords = PS_Script.Streams.Error.ReadAll();
                _WarningRecords = PS_Script.Streams.Warning.ReadAll();
                _InformationRecords = PS_Script.Streams.Information.ReadAll();
                if (_ErrorRecords != null && _ErrorRecords.Count >= 1)
                    _Status = "Error";
                else if (_WarningRecords != null && _WarningRecords.Count >= 1)
                    _Status = "Warning";
                else if (_InformationRecords != null && _InformationRecords.Count >= 1)
                    _Status = "Information";
                else
                    _Status = "Success";
            }
            catch (Exception)
            {
                _ErrorRecords = PS_Script.Streams.Error.ReadAll();
                _WarningRecords = PS_Script.Streams.Warning.ReadAll();
                _InformationRecords = PS_Script.Streams.Information.ReadAll();
                _Status = "Error";
                _Output = null;
            }
        }

        public string GetErrorMessage()
        {
            string ErrorMessage = null;
            if ( _ErrorRecords.Count > 0 )
            {
                foreach (ErrorRecord ERR in _ErrorRecords)
                {
                    ErrorMessage += ERR.Exception.Message;
                    ErrorMessage += "\n*******************\n";
                }
            }
            return ErrorMessage;
        }

        public string GetWarningMessage()
        {
            string WarningMessage = null;
            if (_WarningRecords.Count > 0)
            {
                foreach (WarningRecord WAR in _WarningRecords)
                {
                    WarningMessage += WAR.Message;
                    WarningMessage += "\n*******************\n";
                }
            }
            return WarningMessage;
        }

        public string GetInformationMessage()
        {
            string InformationMessage = null;
            if (_InformationRecords.Count > 0)
            {
                foreach (InformationRecord INFO in _InformationRecords)
                {
                    InformationMessage += INFO.MessageData;
                    InformationMessage += "\n*******************\n";
                }
            }
            return InformationMessage;
        }
    }
}
