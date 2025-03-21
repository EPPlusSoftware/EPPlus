/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.IO;

namespace OfficeOpenXml.Drawing.OleObject
{
    /// <summary>
    /// Object containing additional parameters for OLE Objects.
    /// </summary>
    public class ExcelOleObjectParameters
    {
        private string _olePath = null;
        internal string OlePath
        {
            get 
            {
                return _olePath;
            }
            set
            {
                _olePath = value;
            }
        }

        /// <summary>
        /// True: File will be linked. False: File will be embedded.
        /// </summary>
        public bool LinkToFile = false;
        /// <summary>
        /// Set to display the object as in icon.
        /// </summary>
        public bool DisplayAsIcon = false;
        /// <summary>
        /// Use to set custom progId.
        /// </summary>
        public string ProgId = null;
        /// <summary>
        /// The icon for the object.
        /// </summary>
        public ExcelImage Icon = null;
    }
}
