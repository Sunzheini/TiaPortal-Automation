using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TiaAutomation.Core.Contracts
{
    public interface ITiaOpennessController
    {
        string CreateInstance();

        string OpenSolution(string solutionPath);

        string CreateITcSysManager();

        string ClientConnect();

        string TempMethodOpeness();

        string BuildSolution();

        string SetTargetNetId(string amsNetId, string defaultAmsNetId);

        string ActivateConfiguration();

        string StartRestartTiaPortal();

        string ReadFromPlc(string nameOfIntVarToRead, string defaultNameOfIntVarToRead);

        string ToggleEnable();

        string ToggleStartStop();

        string CloseSolution();

        string KillInstance();

        string ClientDisconnect();
    }
}
