/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControlList : ExcelControl
    {
        internal ExcelControlList(ExcelDrawings drawings, XmlElement drawNode, string name) : base(drawings, drawNode, name)
        {
        }

        internal ExcelControlList(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml)
            : base(drawings, drawNode, control, part, controlPropertiesXml)
        {
        }
        public ExcelAddressBase InputRange 
        { 
            get
            {
                var range = _ctrlProp.GetXmlNodeString("@fmlaRange");
                if(ExcelAddressBase.IsValidAddress(range))
                {
                    return new ExcelAddressBase(range);
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    _ctrlProp.DeleteNode("@fmlaRange");
                    _vmlProp.DeleteNode("x:FmlaRange");
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaRange", value.Address);
                    _vmlProp.SetXmlNodeString("x:FmlaRange", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaRange", value.FullAddress);
                    _vmlProp.SetXmlNodeString("x:FmlaRange", value.FullAddress);
                }
            }
        }
        /// <summary>
        /// The index of a selected item in the list. 
        /// </summary>
        public int SelectedIndex
        {
            get
            {
                return _ctrlProp.GetXmlNodeInt("@sel", 0) - 1;
            }
            set
            {
                if (value <= 0)
                {
                    _ctrlProp.DeleteNode("@sel");
                    _vmlProp.DeleteNode("x:Sel");
                }
                else
                {
                    _ctrlProp.SetXmlNodeInt("@sel", value);
                    _vmlProp.SetXmlNodeInt("x:Sel", value);
                }
            }
        }
        internal int Page
        {
            get
            {
                return _vmlProp.GetXmlNodeInt("x:Page");
            }
            set
            {
                _ctrlProp.SetXmlNodeInt("@page", value);
                _vmlProp.SetXmlNodeInt("x:Page", value);
            }
        }
    }
}