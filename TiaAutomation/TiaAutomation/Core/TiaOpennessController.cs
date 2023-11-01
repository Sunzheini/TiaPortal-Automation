using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Connection;
using Siemens.Engineering.Download;
using Siemens.Engineering.Download.Configurations;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using TiaAutomation.Core.Contracts;
using TiaAutomation.IO.Contracts;
using TiaTest;

using Sharp7;
using static Sharp7.S7Client;
using static Sharp7.S7;
using Siemens.Engineering.Online;

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
        // Tia Openess
        // -------------------

        public static TiaPortal instTIA;   
        public bool openedWithInterface;
        public static Project currentProject;   
        public static ProjectComposition projectsObject;

        public string locationOfTiaExe = "C:\\Program Files\\Siemens\\Automation\\Portal V18\\Bin\\Siemens.Automation.Portal.exe";
        public string projectPath = "C:\\Appl\\Projects\\Siemens\\FBW_Siemens_PLC_OPC_UA_V18\\FBW_Siemens_PLC_OPC_UA_V18.ap18";
        public static FileInfo projectFileInfoPath = null;

        public string projectDirectoryPath = "C:\\Appl\\Projects\\Siemens\\FBW_Siemens_PLC_OPC_UA_V18";
        public static DirectoryInfo projectDirectoryInfoPath = null;

        public static DeviceComposition projectDevicesObject;
        public static Device device;
        public static DeviceItem deviceItem;

        // -------------------

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
            openedWithInterface = true;

            return _returnStringWhenSuccess;
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

            // -------------------

            projectDirectoryInfoPath = new DirectoryInfo(projectDirectoryPath);

            projectFileInfoPath = new FileInfo(projectPath);
            projectsObject = instTIA.Projects;
            currentProject = projectsObject.Open(projectFileInfoPath);

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

            // -------------------
            projectDevicesObject = currentProject.Devices;

            //string result = string.Empty;
            //foreach(Device device in projectDevicesObject)
            //{
            //    result += device.Name + "\n";
            //}
            //return result;

            device = projectDevicesObject[0];

            DeviceItemComposition currentDeviceItemAggregation = device.DeviceItems;

            // [0] is Rail_0, [1] is PLC_1
            deviceItem = currentDeviceItemAggregation[1];

            // return $"Device name: {device.Name}";
            // return $"Device name: {deviceItem.Name}";   // PLC_1

            // -------------------

            return _returnStringWhenSuccess;
        }

        public string ClientConnect()
        {
            return _returnStringWhenSuccess;
        }

        // Move to open solution after security problem is solved
        public string TempMethodOpeness()
        {
            projectFileInfoPath = new FileInfo(projectPath);
            projectsObject = instTIA.Projects;
            currentProject = projectsObject.Open(projectFileInfoPath);

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

            // -------------------

            // container = ((IEngineeringServiceProvider)device).GetService<SoftwareContainer>();
            // softwareBase = container.Software;

            // DeviceItem deviceItemToGetService = deviceItem as DeviceItem;
            // softwareContainer = deviceItemToGetService.GetService<SoftwareContainer>();
            SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();
            PlcSoftware software = softwareContainer.Software as PlcSoftware;

            // get the string repr of the software
            // string softwareString = software.ToString();

            ICompilable compileService = software.GetService<ICompilable>();
            CompilerResult result = compileService.Compile();

            return $"Compiled: {result}";

            // return _returnStringWhenSuccess;
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

            // get first profinet interface
            DeviceItem plcProfinet = deviceItem.DeviceItems.First();
            NetworkInterface plcNetInterface = ((IEngineeringServiceProvider)plcProfinet).GetService<NetworkInterface>();

            if (plcNetInterface != null)
            {
                foreach (Node node in plcNetInterface.Nodes)
                {
                    Console.WriteLine(node.Name);

                    if (node != null)
                    {
                        foreach (EngineeringAttributeInfo nodeInfo in node.GetAttributeInfos())
                        {
                            Console.WriteLine(nodeInfo.Name);

                            if (nodeInfo != null && nodeInfo.Name == "Address")
                            {
                                node.SetAttribute("Address", "192.168.11.11");
                            }
                        }
                    }
                }
            }

            return _returnStringWhenSuccess;
        }

        // --------------------------------------------------------------------------------------------------
        // copied
        // --------------------------------------------------------------------------------------------------
        /// <summary>
        /// Downloads the TIA Portal project on the plc
        /// </summary>
        /// <param name="device">PLC device</param>
        /// <param name="configurationTargetInterface">Configuration of the plc interface</param>
        private static void Download(Device device, ConfigurationTargetInterface configurationTargetInterface)
        {
            DownloadConfigurationDelegate preDownloadDelegate = PreConfigureDownload;
            DownloadConfigurationDelegate postDownloadDelegate = PostConfigureDownload;
            DownloadProvider downloadProvider = null;

            foreach (var item in device.DeviceItems[0].DeviceItems)
            {
                downloadProvider = item.GetService<DownloadProvider>();
                {
                    if (downloadProvider != null)
                    {
                        break;
                    }
                }
            }

            downloadProvider.Configuration.ApplyConfiguration(configurationTargetInterface);
            IConfiguration configuration = configurationTargetInterface;
            downloadProvider.Download(configuration, preDownloadDelegate, postDownloadDelegate, DownloadOptions.Hardware | DownloadOptions.Software);
        }

        /// <summary>
        /// Configurates the download before the download starts
        /// </summary>
        /// <param name="downloadConfiguration">Configuration of the download</param>
        private static void PreConfigureDownload(DownloadConfiguration downloadConfiguration)
        {
            StopModules stopModules = downloadConfiguration as StopModules;
            if (stopModules != null)
            {
                stopModules.CurrentSelection = StopModulesSelections.StopAll; // This selection will set PLC into "Stop" mode
                return;
            }

            OverwriteSystemData overwriteSystemData = downloadConfiguration as OverwriteSystemData;
            if (overwriteSystemData != null)
            {
                overwriteSystemData.CurrentSelection = OverwriteSystemDataSelections.Overwrite;
                return;
            }

            ActiveTestCanBeAborted activeTestCanBeAborted = downloadConfiguration as ActiveTestCanBeAborted;
            if (activeTestCanBeAborted != null)
            {
                activeTestCanBeAborted.CurrentSelection = ActiveTestCanBeAbortedSelections.AcceptAll;
                return;
            }

            AlarmTextLibrariesDownload alarmTextLibraries = downloadConfiguration as AlarmTextLibrariesDownload;

            if (alarmTextLibraries != null)
            {
                alarmTextLibraries.CurrentSelection = AlarmTextLibrariesDownloadSelections.ConsistentDownload;
                return;
            }

            BlockBindingPassword blockBindingPassword = downloadConfiguration as BlockBindingPassword;

            if (blockBindingPassword != null)
            {
                SecureString password = null; // Get Binding password from a secure location
                blockBindingPassword.SetPassword(password);
                return;
            }
            CheckBeforeDownload checkBeforeDownload = downloadConfiguration as CheckBeforeDownload;

            if (checkBeforeDownload != null)
            {
                checkBeforeDownload.Checked = true;
                return;
            }

            ConsistentBlocksDownload consistentBlocksDownload = downloadConfiguration as ConsistentBlocksDownload;

            if (consistentBlocksDownload != null)
            {
                consistentBlocksDownload.CurrentSelection = ConsistentBlocksDownloadSelections.ConsistentDownload;
                return;
            }

            ModuleWriteAccessPassword moduleWriteAccessPassword = downloadConfiguration as ModuleWriteAccessPassword;

            if (moduleWriteAccessPassword != null)
            {
                SecureString password = null; // Get PLC protection level password from a secure location

                moduleWriteAccessPassword.SetPassword(password);
                return;
            }
            throw new NotSupportedException(); // Exception thrown in the delagate will cancel download
        }

        /// <summary>
        /// Configurates the download after the download finished
        /// </summary>
        /// <param name="downloadConfiguration">Configuration of the download</param>
        private static void PostConfigureDownload(DownloadConfiguration downloadConfiguration)
        {
            StartModules startModules = downloadConfiguration as StartModules;
            if (startModules != null)
            {
                startModules.CurrentSelection = StartModulesSelections.StartModule;
                return;
            }
        }

        /// <summary>
        /// Writes the result of the download process in the command line
        /// </summary>
        /// <param name="result">Result of the download process</param>
        private static void WriteDownloadResults(DownloadResult result)
        {
            Console.WriteLine("State:" + result.State);
            Console.WriteLine("Warning Count:" + result.WarningCount);
            Console.WriteLine("Error Count:" + result.ErrorCount);
            RecursivelyWriteMessages(result.Messages);
        }

        /// <summary>
        /// Writes all results of the download process in the command line
        /// </summary>
        /// <param name="messages">Result messages of the download process</param>
        /// <param name="indent">Indent settings</param>
        private static void RecursivelyWriteMessages(DownloadResultMessageComposition messages,
        string indent = "")
        {
            indent += "\t";
            foreach (DownloadResultMessage message in messages)
            {
                Console.WriteLine(indent + "DateTime: " + message.DateTime);
                Console.WriteLine(indent + "State: " + message.State);
                Console.WriteLine(indent + "Message: " + message.Message);
                RecursivelyWriteMessages(message.Messages, indent);
            }
        }

        // --------------------------------------------------------------------------------------------------
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
            // page 482
            const string networkAdapter = @"ASIX AX88179 USB 3.0 to Gigabit Ethernet Adapter";
            const string configurationMode = @"PN/IE";

            DownloadProvider downloadProvider = deviceItem.GetService<DownloadProvider>();
            ConnectionConfiguration configuration = downloadProvider.Configuration;
            ConfigurationMode mode = configuration.Modes.Find(configurationMode);
            ConfigurationPcInterface pcInterface = mode.PcInterfaces.Find(networkAdapter, 1);

            IConfiguration targetConfiguration = pcInterface.TargetInterfaces[0];

            DownloadConfigurationDelegate preDownloadDelegate = PreConfigureDownload;
            DownloadConfigurationDelegate postDownloadDelegate = PostConfigureDownload;

            try
            {
                DownloadResult result = downloadProvider.Download(targetConfiguration, preDownloadDelegate, postDownloadDelegate, DownloadOptions.Hardware | DownloadOptions.Software);
                // DownloadResult result = downloadProvider.Download(targetConfiguration, preDownloadDelegate, postDownloadDelegate, DownloadOptions.Software);
                return $"Download result: {result}";
            }
            catch (Exception e)
            {
                return $"{e}";
            }

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

            // ----------------------------------------------------
            // Sharp7
            // ----------------------------------------------------
            string ipAddress = "192.168.2.12";

            var client = new S7Client();

            // public int ConnectTo(String Address, int Rack, int Slot)
            // try 0, 0 or 0, 1
            int connectionResult = client.ConnectTo(ipAddress, 0, 0);

            // 0 : The Client is successfully connected (or was already connected).
            if (connectionResult == 0)
            {
                // start the plc (also has PlcHotStart)
                var resultFromStart = client.PlcColdStart();

                //check if started
                if (resultFromStart == 0)
                {
                    // wait for 10 seconds
                    System.Threading.Thread.Sleep(10000);

                    // stop the plc
                    var resultFromStop = client.PlcStop();

                    // check if stopped
                    if (resultFromStop == 0)
                    {
                        // disconnect from plc
                        client.Disconnect();
                    }
                    else
                    {
                        return "Error while stopping PLC";
                    }
                }
                else
                {
                    return "Error while starting PLC";
                }
            }
            else
            {
                return "Error while connecting to PLC";
            }

            // ----------------------------------------------------
            // Openness
            // ----------------------------------------------------

            return _returnStringWhenSuccess;
        }

        public string GetCpuInfo(S7Client client)
        {
            var cpuInfo = new S7CpuInfo();
            int cpuStatus = 0;

            client.GetCpuInfo(ref cpuInfo);
            Console.WriteLine("##################################");
            Console.WriteLine("#############CPU-Info#############");
            Console.WriteLine("##################################");
            Console.WriteLine("Module Name: " + cpuInfo.ModuleName);
            Console.WriteLine("AS Name: " + cpuInfo.ASName);
            Console.WriteLine("Module Type Name: " + cpuInfo.ModuleTypeName);
            Console.WriteLine("Serialnumber: " + cpuInfo.SerialNumber);
            Console.WriteLine("##################################");

            client.PlcGetStatus(ref cpuStatus);
            Console.WriteLine("############CPU-Status############");
            Console.WriteLine("##################################");
            Console.WriteLine("Status: " + cpuStatus);

            return _returnStringWhenSuccess;
        }

        public string GoOnline()
        {
            const string networkAdapter = @"ASIX AX88179 USB 3.0 to Gigabit Ethernet Adapter";
            const string configurationMode = @"PN/IE";
            const string slotString = @"2 X3";

            OnlineProvider onlineProvider = deviceItem.GetService<OnlineProvider>();
            ConnectionConfiguration configuration = onlineProvider.Configuration;
            ConfigurationMode mode = configuration.Modes.Find(configurationMode);
            ConfigurationPcInterface pcInterface = mode.PcInterfaces.Find(networkAdapter, 1);
            ConfigurationTargetInterface slot = pcInterface.TargetInterfaces.Find(slotString);
            configuration.ApplyConfiguration(slot);
            onlineProvider.GoOnline();

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

            currentProject.Close();
            
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

            // -------------------

            if (openedWithInterface)
            {
                // Get the process name without the path.
                string tiaExeName = System.IO.Path.GetFileNameWithoutExtension(locationOfTiaExe);

                // Find the TIA Portal process by its name.
                Process[] processes = Process.GetProcessesByName(tiaExeName);

                // Terminate all instances of the TIA Portal process.
                foreach (Process process in processes)
                {
                    process.Kill();
                }
            }
            else    // if without interface
            {
                instTIA.Dispose();
            }

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
