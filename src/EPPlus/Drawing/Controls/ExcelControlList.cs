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
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public abstract class ExcelControlList : ExcelControl
    {
        internal ExcelControlList(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackageRelationship rel, XmlDocument controlPropertiesXml)
            : base(drawings, drawNode, control, rel,  controlPropertiesXml, null)
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
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaRange", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaRange", value.FullAddress);
                }
            }
        }
        public ExcelAddressBase LinkedCell
        {
            get
            {
                var range = _ctrlProp.GetXmlNodeString("@fmlaLink");
                if (ExcelAddressBase.IsValidAddress(range))
                {
                    return new ExcelAddressBase(range);
                }
                return null;
            }
            set
            {
                if(value == null)
                {
                    _ctrlProp.DeleteNode("@fmlaLink");
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaLink", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaLink", value.FullAddress);
                }
            }
        }
    }
}