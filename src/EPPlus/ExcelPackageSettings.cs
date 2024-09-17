/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/10/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System;
using OfficeOpenXml.Interfaces.SensitivityLabels;
using OfficeOpenXml.SensitivityLabels;

namespace OfficeOpenXml
{
    /// <summary>
    /// Package generic settings
    /// </summary>
    public class ExcelPackageSettings
    {
        /// <summary>
        /// Culture specific number formats for the build-in number formats ranging from 14-47.
        /// As some build-in number formats are culture specific, this collection adds the pi
        /// </summary>
        public static Dictionary<string, Dictionary<int, string>> CultureSpecificBuildInNumberFormats
        {
            get;
        } = new Dictionary<string, Dictionary<int, string>>(StringComparer.InvariantCultureIgnoreCase);

        internal ExcelPackageSettings()
        {

        }
        /// <summary>
        /// Do not call garbage collection when ExcelPackage is disposed.
        /// </summary>
        public bool DoGarbageCollectOnDispose { get; set; } = true;
        
        private ExcelTextSettings _textSettings = null;
        /// <summary>
        /// Manage text settings such as measurement of text for the Autofit functions.
        /// </summary>
        public ExcelTextSettings TextSettings
        {
            get
            {
                if (_textSettings == null)
                {                    
                    _textSettings = new ExcelTextSettings();
                }
                return _textSettings;
            }
        }
        private ExcelImageSettings _imageSettings = null;
        /// <summary>
        /// Set the handler for getting image bounds. 
        /// </summary>
        public ExcelImageSettings ImageSettings
        {
            get
            {
                if (_imageSettings == null)
                {
                    _imageSettings = new ExcelImageSettings();
                }
                return _imageSettings;
            }
        }
        /// <summary>
        /// Any auto- or table- filters created will be applied on save.
        /// In the case you want to handle this manually, set this property to false.
        /// </summary>
        public bool ApplyFiltersOnSave { get; set; } = true;
#if(!NET35)
        ISensitivityLabelHandler _sensibilityLabelHandler = null;
        /// <summary>
        /// If you want your workbooks to be marked with sensibility lables, you can add a handler for authentication, encryption and decryption using the Microsoft Information Protection SDK.
        /// For more information
        /// </summary>
        public ISensitivityLabelHandler SensibilityLabelHandler
        {
            get
            {
                if(_sensibilityLabelHandler==null)
                {
                    return ExcelSensibilityLabels.SensibilityLabelHandler;
                }
                return _sensibilityLabelHandler;
            }
            set
            {
                _sensibilityLabelHandler = value;
            }
        }
    }
#endif
}
