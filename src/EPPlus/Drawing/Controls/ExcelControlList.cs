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
        public ExcelAddressBase LinkedCell
        {
            get
            {
                var range = _ctrlProp.GetXmlNodeString("@fmlaGroup");
                if (string.IsNullOrEmpty(range))
                {
                    range = _ctrlProp.GetXmlNodeString("@fmlaLink");
                }
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
                    _vmlProp.DeleteNode("x:FmlaLink");
                }
                if (value.WorkSheetName.Equals(_drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    _ctrlProp.SetXmlNodeString("@fmlaLink", value.Address);
                    _vmlProp.SetXmlNodeString("x:FmlaLink", value.Address);
                }
                else
                {
                    _ctrlProp.SetXmlNodeString("@fmlaLink", value.FullAddress);
                    _vmlProp.SetXmlNodeString("x:FmlaLink", value.FullAddress);
                }
            }
        }
        /// <summary>
        /// The type of selection
        /// </summary>
        public eSelectionType SelectionType 
        { 
            get
            {
                return _ctrlProp.GetXmlNodeString("@selType").ToEnum(eSelectionType.Single);
            }
            set
            {
                _ctrlProp.SetXmlNodeString("selType", value.ToEnumString());
                _vmlProp.SetXmlNodeString("x:SelType", value.ToString());
            }
        }
        /// <summary>
        /// If <see cref="SelectionType"/> is Multi or extended this array contains the selected indicies. Index is zero based. 
        /// </summary>
        public int[] MultiSelection
        {
            get
            {
                var s=_ctrlProp.GetXmlNodeString("@multiSel");
                if(string.IsNullOrEmpty(s))
                {
                    return null;
                }
                else
                {
                    var a = s.Split(',');
                    try
                    {
                        return a.Select(x => int.Parse(x)-1).ToArray();
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
            set
            {
                if(value==null)
                {
                    _ctrlProp.DeleteNode("@multiSel");
                    _vmlProp.DeleteNode("x:MultiSel");
                }
                var v = value.Select(x => (x+1).ToString(CultureInfo.InvariantCulture)).Aggregate((x, y) => x + "," + y);
                _ctrlProp.SetXmlNodeString("selType", v);
                _vmlProp.SetXmlNodeString("x:MultiSel", v);
            }
        }
        /// <summary>
        /// The index of a selected item in the list. 
        /// </summary>
        public int SelectedIndex
        {
            get
            {
                return GetXmlNodeInt("@sel", 0) - 1;
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
    }
}