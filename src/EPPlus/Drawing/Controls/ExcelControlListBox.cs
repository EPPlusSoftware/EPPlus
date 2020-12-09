/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    10/21/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    public class ExcelControlListBox : ExcelControlList
    {
        internal ExcelControlListBox(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
        {
        }
        internal ExcelControlListBox(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
            : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
        {
        }

        public override eControlType ControlType => eControlType.ListBox;
        /// <summary>
        /// The type of selection
        /// </summary>
        public eSelectionType SelectionType
        {
            get
            {
                return _ctrlProp.GetXmlNodeString("@seltype").ToEnum(eSelectionType.Single);
            }
            set
            {
                _ctrlProp.SetXmlNodeString("@seltype", value.ToEnumString());
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
                var s = _ctrlProp.GetXmlNodeString("@multiSel");
                if (string.IsNullOrEmpty(s))
                {
                    return null;
                }
                else
                {
                    var a = s.Split(',');
                    try
                    {
                        return a.Select(x => int.Parse(x) - 1).ToArray();
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
            set
            {
                if (value == null)
                {
                    _ctrlProp.DeleteNode("@multiSel");
                    _vmlProp.DeleteNode("x:MultiSel");
                }
                var v = value.Select(x => (x + 1).ToString(CultureInfo.InvariantCulture)).Aggregate((x, y) => x + "," + y);
                _ctrlProp.SetXmlNodeString("selType", v);
                _vmlProp.SetXmlNodeString("x:MultiSel", v);
            }
        }
        internal override void UpdateXml()
        {
            base.UpdateXml();
            ((ExcelControlList)this).Page = (int)Math.Round((_height / 14));
        }
    }
}