/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/19/2022         EPPlus Software AB       Implementing handling of initialization errors in ExcelPackage class.
 *************************************************************************************************/
#if (Core)
using Microsoft.Extensions.Configuration;
#else
using System.Configuration;
#endif
using OfficeOpenXml.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    internal static class ExcelConfigurationReader
    {
        /// <summary>
        /// Reads an environment variable from the o/s. If an error occors it will rethrow the <see cref="Exception"/> unless SuppressInitializationExceptions of the <paramref name="config"/> is set to true.
        /// </summary>
        /// <param name="key">The key of the requested variable</param>
        /// <param name="target">The <see cref="EnvironmentVariableTarget"/></param>
        /// <param name="config">Configuration of the package</param>
        /// <param name="initErrors">A list of logged <see cref="ExcelInitializationError"/> objects.</param>
        /// <returns>The value of the environment variable</returns>
        internal static string GetEnvironmentVariable(string key, EnvironmentVariableTarget target, ExcelPackageConfiguration config, List<ExcelInitializationError> initErrors)
        {
            var supressInitExceptions = config.SuppressInitializationExceptions;
            try
            {
                return Environment.GetEnvironmentVariable(key, target);
            }
            catch (Exception ex)
            {
                if (supressInitExceptions)
                {
                    var errorMessage = $"Could not read environment variable \"{key}\"";
                    var error = new ExcelInitializationError(errorMessage, ex);
                    initErrors.Add(error);
                }
                else
                {
                    throw;
                }
            }
            return default;
        }

#if (Core)
        internal static string GetJsonConfigValue(string key, ExcelPackageConfiguration config, List<ExcelInitializationError> initErrors)
        {
            var supressInitExceptions = config.SuppressInitializationExceptions;
            var basePath = config.JsonConfigBasePath;
            var configFileName = config.JsonConfigFileName;
            var configRoot = default(IConfigurationRoot);
            try
            {
                
                var build = new ConfigurationBuilder()
                       .SetBasePath(basePath)
                       .AddJsonFile(configFileName, true, false);
                configRoot = build.Build();
            }
            catch (Exception ex)
            {
                if (supressInitExceptions)
                {
                    var errorMessage = $"Could not load configuration file \"{configFileName}\"";
                    var error = new ExcelInitializationError(errorMessage, ex);
                    initErrors.Add(error);
                }
                else
                {
                    throw;
                }
            }
            if (configRoot != null)
            {
                try
                {
                    var v = configRoot[key];
                    return v;
                }
                catch (Exception ex)
                {
                    if (supressInitExceptions)
                    {
                        var errorMessage = $"Could read key \"{key}\" from appsettings.json";
                        var error = new ExcelInitializationError(errorMessage, ex);
                        initErrors.Add(error);
                        return null;
                    }
                    throw;
                }
            }
            return null;
        }
#endif

#if (!Core)
        internal static string GetValueFromAppSettings(string key, ExcelPackageConfiguration config, List<ExcelInitializationError> initErrors)
        {
            var supressInitExceptions = config.SuppressInitializationExceptions;
            try
            {
                return ConfigurationManager.AppSettings[key];
            }
            catch(Exception ex)
            {
                if (supressInitExceptions)
                {
                    var errorMessage = $"Could read key \"{key}\" from ConfigurationManager.AppSettings";
                    var error = new ExcelInitializationError(errorMessage, ex);
                    initErrors.Add(error);
                    return null;
                }
                throw;
            }
        }
#endif
    }
}
