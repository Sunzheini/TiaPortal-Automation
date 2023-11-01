using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TiaAutomation.Core.Contracts;
using TiaAutomation.IO;
using TiaAutomation.IO.Contracts;
using System.Runtime.Remoting.Messaging;
using TiaAutomation.Core;

namespace TiaAutomation.Core
{
    public class Engine : IEngine
    {
        // this is needed for correct work of the program
        private string _productId = "TcXaeShell.DTE.15.0";

        // default path of the solution (also has a setter in order to change it through a browse button)
        private string _solutionPath = "C:\\Appl\\Projects\\TwinCAT\\Test_Counter\\Test_Counter.sln";

        // either config file or inputs from the front end
        private string _defaultAmsNetId = "5.29.223.252.1.1"; // PLC
        private string _defaultNameOfIntVarToRead = "MAIN.uiCounter";
        private string _nameOfEnableVar = "MAIN.boEnable";
        private int _portForAds = 851;

        //strings for return values
        private string _returnStringWhenSuccess = "Success";
        private string _returnStringStart = "Starting Finished";
        private string _returnStringBuild = "Build Finished";
        private string _returnStringSetTargetNetId = "Set TargetNetID Finished";
        private string _returnStringActivateConfiguration = "Configuration Activated";
        private string _returnStringStartRestartTwinCAT = "TwinCAT (re)started";
        private string _returnStringExit = "Exited";

        private IReader reader;
        private IWriter writer;
        private ITiaOpennessController tiaOpennessController;

        private string _resultStringFromMethodExecution = string.Empty;

        public Engine()
        {
            reader = new Reader();
            writer = new Writer();
            tiaOpennessController = new TiaOpennessController
            (
                _productId
            );
        }

        public string SolutionPath
        {
            get { return _solutionPath; }
            set { _solutionPath = value; }
        }

        public IWriter GetWriter()
        {
            return writer;
        }

        public IReader GetReader()
        {
            return reader;
        }

        // catch the exception here to reduce the code
        public string Start()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.CreateInstance();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // method #2
            _resultStringFromMethodExecution = tiaOpennessController.OpenSolution(_solutionPath);
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // method #3
            _resultStringFromMethodExecution = tiaOpennessController.CreateITcSysManager();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // method #4
            _resultStringFromMethodExecution = tiaOpennessController.ClientConnect();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _returnStringStart;
        }

        public string TempMethod()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.TempMethodOpeness();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _returnStringBuild;
        }

        public string BuildSolution()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.BuildSolution();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _returnStringBuild;
        }

        public string SetTargetNetId(string amsNetId)
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.SetTargetNetId(amsNetId, _defaultAmsNetId);
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _returnStringSetTargetNetId;
        }

        public string ActivateConfiguration()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.ActivateConfiguration();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _returnStringActivateConfiguration;
        }

        public string StartRestartTiaPortal()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.StartRestartTiaPortal();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _returnStringStartRestartTwinCAT;
        }

        public string ReadFromPlc(string nameOfIntVarToRead)
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.ReadFromPlc(nameOfIntVarToRead, _defaultNameOfIntVarToRead);
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _resultStringFromMethodExecution;
        }

        public string ToggleEnableDisable()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.ToggleEnable();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _resultStringFromMethodExecution;
        }

        public string ToggleStartStop()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.ToggleStartStop();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _resultStringFromMethodExecution;
        }

        public string Exit()
        {
            // method #1
            _resultStringFromMethodExecution = tiaOpennessController.CloseSolution();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // method #2
            _resultStringFromMethodExecution = tiaOpennessController.KillInstance();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // method #3
            _resultStringFromMethodExecution = tiaOpennessController.ClientDisconnect();
            if (_resultStringFromMethodExecution != _returnStringWhenSuccess) return _resultStringFromMethodExecution;

            // end
            return _returnStringExit;
        }
    }
}
