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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Configuration
{
    /// <summary>
    /// Parameters for configuring the <see cref="ExcelPackage"/> class before usage
    /// </summary>
    public class ExcelPackageConfiguration
    {
        /// <summary>
        /// If set to true errors/exceptions that occurs during initialization of the ExcelPackage class will
        /// be suppressed and logged in <see cref="ExcelPackage.InitializationErrors"/>.
        /// 
        /// If set to false these Exceptions will be rethrown.
        /// 
        /// Default value of this property is false.
        /// </summary>
        public bool SuppressInitializationExceptions
        {
            get; set;
        }

        private string _jsonConfigBasePath = Directory.GetCurrentDirectory();
        /// <summary>
        /// Path of the directory where the json configuration file is located.
        /// Default value is the path returned from <see cref="System.IO.Directory.GetCurrentDirectory"/>
        /// </summary>
        public string JsonConfigBasePath
        {
            get { return _jsonConfigBasePath; }
            set { _jsonConfigBasePath = value; }
        }

        private string _jsonConfigFileName = "appsettings.json";
        /// <summary>
        /// File name of the json configuration file.
        /// Default value is appsettings.json
        /// </summary>
        public string JsonConfigFileName
        {
            get { return _jsonConfigFileName; }
            set { _jsonConfigFileName = value; }
        }

        /// <summary>
        /// Configuration with default values.
        /// </summary>
        internal static ExcelPackageConfiguration Default
        {
            get { return new ExcelPackageConfiguration(); }
        }

        internal void CopyFrom(ExcelPackageConfiguration other)
        {
            _jsonConfigBasePath = other.JsonConfigBasePath;
            _jsonConfigFileName = other.JsonConfigFileName;
            SuppressInitializationExceptions = other.SuppressInitializationExceptions;
        }

        /// <summary>
        /// Resets configuration to its default values
        /// </summary>
        public void Reset()
        {
            CopyFrom(Default);
        }
    }
}
