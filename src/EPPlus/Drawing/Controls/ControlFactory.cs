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
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    internal static class ControlFactory
    {
        /*
     *<xsd:simpleType name="ST_ObjectType">
   <xsd:restriction base="xsd:token">
     <xsd:enumeration value="Button"/>
     <xsd:enumeration value="CheckBox"/>
     <xsd:enumeration value="Drop"/>
     <xsd:enumeration value="GBox"/>
     <xsd:enumeration value="Label"/>
     <xsd:enumeration value="List"/>
     <xsd:enumeration value="Radio"/>
     <xsd:enumeration value="Scroll"/>
     <xsd:enumeration value="Spin"/>
     <xsd:enumeration value="EditBox"/>
     <xsd:enumeration value="Dialog"/>
   </xsd:restriction>
 </xsd:simpleType> 
     */
        private static Dictionary<string, eControlType> _controlTypeMapping = new Dictionary<string, eControlType>
        {
            { "Drop", eControlType.DropDown },
            { "GBox", eControlType.GroupBox },
            { "List", eControlType.ListBox },
            { "Radio", eControlType.RadioButton },
            { "Scroll", eControlType.ScrollBar },
            { "Spin", eControlType.SpinButton }
        };

        private static eControlType GetControlType(string input)
        {
            if(_controlTypeMapping.ContainsKey(input))
            {
                return _controlTypeMapping[input];
            }
            else
            {
                return input.ToEnum(eControlType.Label);
            }
        }

        internal static ExcelDrawing GetControl(ExcelDrawings drawings, XmlElement drawNode, ControlInternal control)
        {
            var rel = drawings.Worksheet.Part.GetRelationship(control.RelationshipId);
            var controlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            var part = drawings.Worksheet._package.ZipPackage.GetPart(controlUri);
            var controlPropertiesXml = new XmlDocument();
            XmlHelper.LoadXmlSafe(controlPropertiesXml, part.GetStream());
            var objectType = controlPropertiesXml.DocumentElement.Attributes["objectType"]?.Value;
            var controlType = GetControlType(objectType);
            switch(controlType)
            {
                case eControlType.Button:
                    return new ExcelControlButton(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.DropDown:
                    return new ExcelControlDropDown(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.GroupBox:
                    return new ExcelControlGroupBox(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.Label:
                    return new ExcelControlLabel(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.ListBox:
                    return new ExcelControlListBox(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.CheckBox:
                    return new ExcelControlCheckBox(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.RadioButton:
                    return new ExcelControlRadioButton(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.ScrollBar:
                    return new ExcelControlScrollBar(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.SpinButton:
                    return new ExcelControlSpinButton(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.EditBox:
                    return new ExcelControlEditBox(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                case eControlType.Dialog:
                    return new ExcelControlDialog(drawings, drawNode.ParentNode, control, part, controlPropertiesXml);
                default:
                    throw new NotSupportedException();
            }
            throw new NotImplementedException();
        }

        internal static ExcelControl CreateControl(eControlType controlType,ExcelDrawings drawings, XmlElement drawNode, string name)
        {
            ExcelControl ctrl;
            switch (controlType)
            {
                case eControlType.Button:
                    ctrl = new ExcelControlButton(drawings, drawNode)
                    {
                        Text = name
                    };                    
                    break;
                //case eControlType.DropDown:
                //    return new ExcelControlDropDown(drawings, drawNode);
                //case eControlType.GroupBox:
                //    return new ExcelControlGroupBox(drawings, drawNode);
                //case eControlType.Label:
                //    return new ExcelControlLabel(drawings, drawNode);
                //case eControlType.ListBox:
                //    return new ExcelControlListBox(drawings, drawNode);
                case eControlType.CheckBox:
                    ctrl = new ExcelControlCheckBox(drawings, drawNode)
                    {
                        Text = name
                    };
                    break;
                //case eControlType.RadioButton:
                //    return new ExcelControlRadioButton(drawings, drawNode);
                //case eControlType.ScrollBar:
                //    return new ExcelControlScrollBar(drawings, drawNode);
                //case eControlType.SpinButton:
                //    return new ExcelControlSpinButton(drawings, drawNode);
                //case eControlType.EditBox:
                //    return new ExcelControlEditBox(drawings, drawNode);
                //case eControlType.Dialog:
                //    return new ExcelControlDialog(drawings, drawNode);
                default:
                    throw new NotSupportedException();
            }
            ctrl.Name = name;
            return ctrl;
        }
    }
}
