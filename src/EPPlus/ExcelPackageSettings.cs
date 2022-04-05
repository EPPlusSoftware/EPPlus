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
namespace OfficeOpenXml
{
    /// <summary>
    /// Package generic settings
    /// </summary>
    public class ExcelPackageSettings
    {
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
    }
}
