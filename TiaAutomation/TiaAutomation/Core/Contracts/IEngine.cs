using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using TiaAutomation.IO.Contracts;

namespace TiaAutomation.Core.Contracts
{
    public interface IEngine
    {
        string SolutionPath { get; set; }

        IWriter GetWriter();

        IReader GetReader();

        string Start();

        string BuildSolution();

        string SetTargetNetId(string amsNetId);

        string ActivateConfiguration();

        string StartRestartTiaPortal();

        string ReadFromPlc(string nameOfIntVarToRead);

        string ToggleEnableDisable();

        string ToggleStartStop();

        string Exit();
    }
}
