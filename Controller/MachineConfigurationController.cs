using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using BootCamperMachineConfiguration.Helper;
using BootCamperMachineConfiguration.Model;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;
using static System.Console;

namespace BootCamperMachineConfiguration.Controller
{
    public class MachineConfigurationController
    {
        #region Private Variables

        private readonly IConfiguration _configuration;
        private static string _icEmail = string.Empty;

        #endregion

        #region Constructor

        public MachineConfigurationController(string icEmail)
        {
            _icEmail = icEmail;
            _configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", true, true)
                .Build();
        }

        #endregion

        #region Update Google Sheet with Machine Details

        /// <summary>
        /// Updating Boot Camper Machine in Google Drive
        /// </summary>
        public void UpdateBootCamperMachineDetails()
        {
            var currentSystemProps = GetSystemProperties();
            var validate = ValidateSystemProperties(currentSystemProps);

            var icList = GetIcsFromGoogleDrive();
            var index = (from ic in icList.Values
                where ic.Any(e => e.ToString() == _icEmail)
                select icList.Values.IndexOf(ic)).FirstOrDefault();
            // Why index + 3?
            // We started with retrieving data from third row. So +2
            // Retrieved index value starts from 0 so if it is 1 then actual value is 1+1
            // Finally it 2 + index + 1
            index = index + 3;

            //Update Enough Specifications
            var enoughSpecsStoredCell = _configuration["EnoughSpecsStoredCell"];
            var spreadSheetId = _configuration["SpreadSheetID"];
            GoogleDriveHelper.UpdateValuesInGoogleDrive(spreadSheetId, enoughSpecsStoredCell + index,
                validate ? "Yes" : "No");

            //Update Actual Machine Specs
            var actualMachineSpecs = new StringBuilder();
            actualMachineSpecs.AppendLine(" - CPU:" + currentSystemProps.ProcessorName);
            actualMachineSpecs.AppendLine(" - CPU Score:" + currentSystemProps.CpuBenchMarkScore);
            actualMachineSpecs.AppendLine(" - Memory:" + currentSystemProps.UsableMemory + "G");
            actualMachineSpecs.AppendLine(" - OS:" + currentSystemProps.OperatingSystemEdition);
            actualMachineSpecs.AppendLine(" - Storage:" + currentSystemProps.StorageSpace + "G");
            actualMachineSpecs.AppendLine(" - Free disk space:" + currentSystemProps.FreeDiskSpace + "G");
            actualMachineSpecs.Append(" - Arch:" + currentSystemProps.Architecture);
            var actualMachineSpecsStoredCell = _configuration["ActualMachineSpecsStoredCell"];
            GoogleDriveHelper.UpdateValuesInGoogleDrive(spreadSheetId, actualMachineSpecsStoredCell + index,
                actualMachineSpecs.ToString());
            WriteLine("Machine Configuration Details Updated Successfully in Google Drive in ICs sheets");
        }

        private bool ValidateSystemProperties(SystemProperties currentSystemProps)
        {
            var expectedAllSystemProps = GetExpectedSystemProperties();
            if (expectedAllSystemProps == null)
            {
                WriteLine("Entered IC/BootCamper Email not found in our list...");
                WriteLine("Press any key to continue...");
                ReadKey(true);
                throw new Exception("Entered IC/BootCamper Details not found...");
            }
            //
            foreach (var expectedSystemProps in expectedAllSystemProps)
            {
                var expectedOperatingSystemList = expectedSystemProps.OperatingSystem.Split("&");
                var expectedOperatingSystem = string.Empty;
                foreach (var operatingSystem in expectedOperatingSystemList)
                {
                    if (!operatingSystem.Contains(currentSystemProps.OperatingSystemName)) continue;
                    expectedOperatingSystem = operatingSystem;
                    break;
                }
                var expectedOsVersion = Regex.Replace(expectedOperatingSystem, @"[^0-9]+", string.Empty);

                var currentOperatingSystem = Regex.Replace(currentSystemProps.OperatingSystemEdition, @"[^0-9-.]+", string.Empty);
                var currentOsVersion = currentOperatingSystem.Split(".")[0];

                if (Convert.ToInt32(currentOsVersion) < Convert.ToInt32(expectedOsVersion))
                    return false;
                if (currentSystemProps.CpuBenchMarkScore < expectedSystemProps.MinimalCpuBenchmarkScore)
                    return false;
                if (currentSystemProps.UsableMemory < expectedSystemProps.Memory)
                    return false;
                if (currentSystemProps.StorageSpace < expectedSystemProps.Storage)
                    return false;
                if (currentSystemProps.FreeDiskSpace < expectedSystemProps.FreeDiskSpace)
                    return false;
                if (!currentSystemProps.Architecture.Contains(expectedSystemProps.Architecture))
                    return false;
            }
            return true;
        }

        #endregion

        #region GetExpectedSystemProperties

        /// <summary>
        /// Get expected system properties for the entered IC/BootCamper 
        /// </summary>
        /// <returns></returns>
        private IEnumerable<OutputProperties> GetExpectedSystemProperties()
        {
            var enteredIcProjectDetails = GetICsProjectList().Where(e => e.IcEmail == _icEmail).ToList();
            if (enteredIcProjectDetails.Count <= 0) return null;
            {
                var spreadSheetId = _configuration["SpreadSheetID"];
                var range = _configuration["ExpectedSystemPropertiesSheet"];
                var response = GoogleDriveHelper.GetValuesFromGoogleDrive(spreadSheetId, range,
                    SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.COLUMNS);
                return (from value in response.Values
                    let projectName = value[0].ToString()
                    where enteredIcProjectDetails.Exists(e => e.Projects.Exists(d => d == projectName))
                    select new OutputProperties
                    {
                        OperatingSystem = value[1]?.ToString(),
                        MinimalCpuBenchmarkScore = Convert.ToInt32(value[2]),
                        Memory = Convert.ToDouble(value[3]),
                        Storage = Convert.ToInt64(value[4]),
                        FreeDiskSpace = Convert.ToInt64(value[5]),
                        Architecture = value[6]?.ToString()
                    }).ToList();
            }
        }

        private IEnumerable<IcProperties> GetICsProjectList()
        {
            var response = GetIcsFromGoogleDrive();
            return response.Values.Where(value => value.Count != 0)
                .Select(value => new IcProperties
                {
                    IcEmail = value[1].ToString(),
                    Projects = new List<string>
                        {
                            value[2].ToString(),
                            value[3].ToString(),
                            value[4].ToString(),
                            value[5].ToString()
                        }.Distinct()
                        .ToList()
                }).ToList();
        }

        private ValueRange GetIcsFromGoogleDrive()
        {
            var range = _configuration["IcsSheet"];
            var spreadSheetId = _configuration["SpreadSheetID"];
            var response = GoogleDriveHelper.GetValuesFromGoogleDrive(spreadSheetId, range);
            return response;
        }

        #endregion

        #region GetSystemProperties 

        private SystemProperties GetSystemProperties()
        {
            var processorName = GetProcessorInformation();
            var drivesInfoList = DriveInfo.GetDrives().Where(e => e.IsReady).ToList();
            var systemProperties = new SystemProperties
            {
                OperatingSystemName = GetOperatingSystemName(),
                OperatingSystemEdition = RuntimeInformation.OSDescription,
                Architecture = RuntimeInformation.OSArchitecture.ToString(),
                ProcessorName = processorName,
                UsableMemory = GetInstalledMemory(),
                StorageSpace = GetGigaBytesFromBytes(drivesInfoList.Sum(e => e.TotalSize)),
                FreeDiskSpace = GetGigaBytesFromBytes(drivesInfoList.Sum(e => e.AvailableFreeSpace))
            };
            var response = GetCpuBenchMarkScores();
            var requiredProcessorName = GetProcessorName(processorName);
            systemProperties.CpuBenchMarkScore =
                response.Where(e => e.CpuType.Contains(requiredProcessorName)).ToList()[0].Score;
            return systemProperties;
        }

        private static string GetProcessorName(string processorName)
        {
            processorName = processorName.Replace("Intel(R) Core(TM)", "");
            var startIndex = processorName.IndexOf("CPU", StringComparison.Ordinal);
            var endIndex = processorName.Length - startIndex - 3;
            processorName = processorName
                .Replace(processorName.Substring(startIndex, processorName.Length - endIndex), "").Trim();
            return processorName;
        }

        private IEnumerable<CpuBenchMarkScoreProperties> GetCpuBenchMarkScores()
        {
            var range = _configuration["CpuScoresSheet"];
            var spreadSheetId = _configuration["SpreadSheetID"];
            var response = GoogleDriveHelper.GetValuesFromGoogleDrive(spreadSheetId, range);
            return response.Values.Select(value => new CpuBenchMarkScoreProperties
            {
                CpuType = value[0].ToString(),
                Score = Convert.ToInt32(value[1])
            }).ToList();
        }

        private static string GetOperatingSystemName()
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return "Windows";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return "MacOS";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return "Ubuntu";
            throw new Exception(
                "Incompatible Operating System. Please use one of the Operating System from Windows,Mac or Linux");
        }

        //Only Works for Windows
        private static string GetProcessorInformation()
        {
            var managementClass = new ManagementClass("win32_processor");
            var managementObjectCollection = managementClass.GetInstances();
            var processorInfo = string.Empty;
            foreach (var mo in managementObjectCollection)
            {
                processorInfo = mo["Name"].ToString();
            }
            return processorInfo;
        }

        //Only work for Windows
        private static double GetInstalledMemory()
        {
            var objectQuery = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
            var results = new ManagementObjectSearcher(objectQuery).Get();
            double res = 0;
            foreach (var result in results)
            {
                res = Math.Round((Convert.ToDouble(result["TotalVisibleMemorySize"]) / (1024 * 1024)), 2);
            }
            return res;
        }

        private static long GetGigaBytesFromBytes(long memorySize) => (memorySize / 1024 / 1024 / 1024);

        #endregion
    }
}
