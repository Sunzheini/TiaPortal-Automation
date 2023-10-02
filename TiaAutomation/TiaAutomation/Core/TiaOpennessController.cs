using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.HW;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using TiaAutomation.Core.Contracts;
using TiaAutomation.IO.Contracts;
using TiaTest;

namespace TiaAutomation.Core
{
    public class TiaOpennessController : ITiaOpennessController
    {
        private string _productId = String.Empty;
        private static Type _typeOfInstance = null;
        //private static EnvDTE.DTE _dte = null;
        //private static EnvDTE.Solution _solution = null;
        //private static EnvDTE.Project _project = null;
        //private static ITcSysManager15 _sysManager = null;

        private string _returnStringWhenSuccess = "Success";

        // -------------------

        // Class1 newTIA = null;

        public static TiaPortal instTIA;
        public static Project projectTia;
        public static ProjectComposition projects;

        public string projectPath = "C:\\Appl\\Projects\\Siemens\\Project1\\Project1.ap18";
        public static FileInfo targetDir = null;
        public static Device plcDevice;

        public TiaOpennessController(string productId)
        {
            this._productId = productId;
        }

        // catch exception in Engine.cs, check if exception is the same!
        public string CreateInstance()
        {
            //try
            //{
            //    _typeOfInstance = System.Type.GetTypeFromProgID(_productId);
            //}
            //catch (Exception e)
            //{
            //    return $"Error in GetTypeFromProgID: {e}";
            //}

            //try
            //{
            //    _dte = (EnvDTE.DTE)System.Activator.CreateInstance(_typeOfInstance);
            //}
            //catch (Exception e)
            //{
            //    return $"Error while creating instance: {e}";
            //}

            //try
            //{
            //    _dte.SuppressUI = false;        // when true, the console will not be shown
            //    _dte.MainWindow.Visible = true; // when false, twincat will not be shown
            //}
            //catch (Exception e)
            //{
            //    return $"Error in window options: {e}";
            //}

            //SetWhiteList(
            //    System.Diagnostics.Process.GetCurrentProcess().ProcessName, 
            //    System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName
            //);

            instTIA = new TiaPortal(TiaPortalMode.WithUserInterface);

            return "Test";

            //return _returnStringWhenSuccess;
        }

        public void SetWhiteList(string ApplicationName, string ApplicationStartUpPath)
        {
            RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey software = null;
            try
            {
                software = key.OpenSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .OpenSubKey("18.0")
                    .OpenSubKey("Whitelist")
                    .OpenSubKey(ApplicationName + ".exe")
                    .OpenSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl);
            }
            catch (Exception)
            {
                software = key.CreateSubKey(@"SOFTWARE\Siemens\Automation\Openness")
                    .CreateSubKey("18.0")
                    .CreateSubKey("Whitelist")
                    .CreateSubKey(ApplicationName + ".exe")
                    .CreateSubKey("Entry", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
            }

            string lastWriteTimeUtcFormatted = String.Empty;
            DateTime lastWriteTimeUtc;
            HashAlgorithm hashAlgorithm = SHA256.Create();
            FileStream stream = File.OpenRead(ApplicationStartUpPath);
            byte[] hash = hashAlgorithm.ComputeHash(stream);

            string convertedHash = Convert.ToBase64String(hash);
            software.SetValue("FileHash", convertedHash);
            lastWriteTimeUtc = new FileInfo(ApplicationStartUpPath).LastWriteTimeUtc;
            
            lastWriteTimeUtcFormatted = lastWriteTimeUtc.ToString(@"yyyy/MM/dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
            software.SetValue("DateModified", lastWriteTimeUtcFormatted);
            software.SetValue("Path", ApplicationStartUpPath);
        }

        public string OpenSolution(string solutionPath)
        {
            //try
            //{
            //    _solution = _dte.Solution;
            //    _solution.Open(@solutionPath);
            //}
            //catch (Exception e)
            //{
            //    return $"Error in OpenSolution: {e}";
            //}

            //targetDir = new FileInfo(projectPath);
            //projectTia = instTIA.Projects.Open(targetDir);

            return _returnStringWhenSuccess;
        }

        public string CreateITcSysManager()
        {
            //try
            //{
            //    _project = _solution.Projects.Item(1);
            //    _sysManager = (ITcSysManager15)_project.Object;
            //}
            //catch (Exception e)
            //{
            //    return $"Error in CreateITcSysManager: {e}";
            //}
            return _returnStringWhenSuccess;
        }

        public string ClientConnect()
        {
            return _returnStringWhenSuccess;
        }

        public string BuildSolution()
        {
            //try
            //{
            //    _solution.SolutionBuild.Build(true);
            //}
            //catch (Exception e)
            //{
            //    return $"Error in BuildSolution: {e}";
            //}

            targetDir = new FileInfo(projectPath);
            // projectTia = instTIA.Projects.Open(targetDir);

            projects = instTIA.Projects;
            projectTia = projects.Open(targetDir);

            return _returnStringWhenSuccess;
        }

        public string SetTargetNetId(string amsNetId, string defaultAmsNetId)
        {
            //try
            //{
            //    _sysManager.SetTargetNetId(amsNetId);
            //}
            //catch (Exception e)
            //{
            //    return $"Error in SetTargetNetId: {e}";
            //}
            //finally
            //{
            //    _sysManager.SetTargetNetId(defaultAmsNetId);
            //}
            return _returnStringWhenSuccess;
        }

        public string ActivateConfiguration()
        {
            //try
            //{
            //    _sysManager.ActivateConfiguration();
            //}
            //catch (Exception e)
            //{
            //    return $"Error in ActivateConfiguration: {e}";
            //}
            return _returnStringWhenSuccess;
        }

        public string StartRestartTiaPortal()
        {
            //try
            //{
            //    _sysManager.StartRestartTwinCAT();
            //}
            //catch (Exception e)
            //{
            //    return $"Error in StartRestartTwinCAT: {e}";
            //}
            return _returnStringWhenSuccess;
        }

        public string ReadFromPlc(string nameOfIntVarToRead, string defaultNameOfIntVarToRead)
        {
            //int value = 0;
            //string resultString = string.Empty;

            //// first try the general info
            //try
            //{
            //    DeviceInfo deviceInfo = _client.ReadDeviceInfo();
            //    Version version = deviceInfo.Version.ConvertToStandard();
            //    resultString += $"Device name: {deviceInfo.Name}\n";
            //    resultString += $"Device version: {version}\n";
            //}
            //catch (Exception e)
            //{
            //    return $"Error while reading device info from PLC: {e}";
            //}

            //// now try to read a variable
            //try
            //{
            //    try
            //    {
            //        value = (int)_client.ReadValue
            //        (
            //            //"MAIN.uiCounter",
            //            nameOfIntVarToRead,
            //            typeof(int)
            //        );
            //        resultString += $"Value of variable {nameOfIntVarToRead}: \n {value}\n";
            //    }
            //    catch (Exception e)
            //    {
            //        value = (int)_client.ReadValue
            //        (
            //            defaultNameOfIntVarToRead,
            //            typeof(int)
            //        );
            //        resultString += $"Value of variable {defaultNameOfIntVarToRead}: \n {value}\n";
            //    }
            //}
            //catch (Exception e)
            //{
            //    return $"Error while reading variable from PLC: {e}";
            //}

            //return resultString;

            return _returnStringWhenSuccess;
        }

        public string ToggleEnable()
        {
            //string resultString = string.Empty;
            //bool valueToWrite = true;

            //try
            //{
            //    var boolStatus = _client.ReadValue(_nameOfEnableVar, typeof(bool));

            //    if (boolStatus is bool)
            //    {
            //        bool isTrue = (bool)boolStatus;

            //        if (isTrue)
            //        {
            //            // Write false to the PLC
            //            valueToWrite = false;
            //            _client.WriteValue(_nameOfEnableVar, valueToWrite);
            //            resultString = $"Value written: {valueToWrite}";
            //        }
            //        else
            //        {
            //            // Write true to the PLC
            //            valueToWrite = true;
            //            _client.WriteValue(_nameOfEnableVar, valueToWrite);
            //            resultString = $"Value written: {valueToWrite}";
            //        }
            //    }
            //    else
            //    {
            //        resultString = "Value read is not boolean!";
            //    }
            //}
            //catch (Exception e)
            //{
            //    return $"Error while enabling / disabling: {e}";
            //}
            //return resultString;

            return _returnStringWhenSuccess;
        }

        public string ToggleStartStop()
        {
            //string resultString = string.Empty;

            //try
            //{
            //    StateInfo stateInfo = _client.ReadState();
            //    AdsState state = AdsState.Invalid;
            //    state = stateInfo.AdsState;
            //    short deviceState = stateInfo.DeviceState;

            //    if (state == AdsState.Stop)
            //    {
            //        _client.WriteControl(new StateInfo(AdsState.Run, 0));
            //    }
            //    else if (state == AdsState.Run)
            //    {
            //        _client.WriteControl(new StateInfo(AdsState.Stop, 0));
            //    }

            //    stateInfo = _client.ReadState();
            //    state = stateInfo.AdsState;
            //    deviceState = stateInfo.DeviceState;

            //    resultString += $"DeviceState: {deviceState}\n";
            //    resultString += $"AdsState: {state}\n";
            //}
            //catch (Exception e)
            //{
            //    return $"Error while starting / stopping: {e}";
            //}
            //return resultString;

            return _returnStringWhenSuccess;
        }

        public string CloseSolution()
        {
            //try
            //{
            //    _solution.SaveAs(_solution.FullName);
            //    _solution.Close();
            //}
            //catch (Exception e)
            //{
            //    return $"Error in CloseSolution: {e}";
            //}
            return _returnStringWhenSuccess;
        }

        public string KillInstance()
        {
            //try
            //{
            //    _dte.Quit();
            //}
            //catch (Exception e)
            //{
            //    return $"Error in KillInstance: {e}";
            //}
            return _returnStringWhenSuccess;
        }

        public string ClientDisconnect()
        {
            //try
            //{
            //    this._client.Disconnect();
            //    this._client.Dispose();
            //}
            //catch (Exception e)
            //{
            //    return $"Error while disconnecting from PLC: {e}";
            //}
            return _returnStringWhenSuccess;
        }
    }
}
