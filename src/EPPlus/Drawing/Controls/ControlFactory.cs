/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/01/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
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

        internal static ExcelDrawing GetControl(ExcelDrawings drawings, XmlElement drawNode, ControlInternal control, ExcelGroupShape parent)
        {
            var rel = drawings.Worksheet.Part.GetRelationship(control.RelationshipId);
            var controlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            var part = drawings.Worksheet._package.ZipPackage.GetPart(controlUri);
            var controlPropertiesXml = new XmlDocument();
            XmlHelper.LoadXmlSafe(controlPropertiesXml, part.GetStream());
            var objectType = controlPropertiesXml.DocumentElement.Attributes["objectType"]?.Value;
            var controlType = GetControlType(objectType);
            
            XmlNode node;            
            if(parent==null)
            {
                node = drawNode.ParentNode;
            }
            else
            {
                node = drawNode;
            }

            switch(controlType)
            {
                case eControlType.Button:
                    return new ExcelControlButton(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.DropDown:
                    return new ExcelControlDropDown(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.GroupBox:
                    return new ExcelControlGroupBox(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.Label:
                    return new ExcelControlLabel(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.ListBox:
                    return new ExcelControlListBox(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.CheckBox:
                    return new ExcelControlCheckBox(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.RadioButton:
                    return new ExcelControlRadioButton(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.ScrollBar:
                    return new ExcelControlScrollBar(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.SpinButton:
                    return new ExcelControlSpinButton(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.EditBox:
                    return new ExcelControlEditBox(drawings, node, control, part, controlPropertiesXml, parent);
                case eControlType.Dialog:
                    return new ExcelControlDialog(drawings, node, control, part, controlPropertiesXml, parent);
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
                    ctrl = new ExcelControlButton(drawings, drawNode, name);
                    break;
                case eControlType.DropDown:
                    ctrl = new ExcelControlDropDown(drawings, drawNode, name);
                    break;
                case eControlType.GroupBox:
                    ctrl = new ExcelControlGroupBox(drawings, drawNode, name);
                    break;
                case eControlType.Label:
                    ctrl = new ExcelControlLabel(drawings, drawNode, name);
                    break;
                case eControlType.ListBox:
                    ctrl = new ExcelControlListBox(drawings, drawNode, name);
                    break;
                case eControlType.CheckBox:
                    ctrl = new ExcelControlCheckBox(drawings, drawNode, name);
                    break;
                case eControlType.RadioButton:
                    ctrl = new ExcelControlRadioButton(drawings, drawNode, name);
                    break;
                case eControlType.ScrollBar:
                    ctrl=new ExcelControlScrollBar(drawings, drawNode, name);
                    break;
                case eControlType.SpinButton:
                    ctrl = new ExcelControlSpinButton(drawings, drawNode, name);
                    break;
                //case eControlType.EditBox:
                //    return new ExcelControlEditBox(drawings, drawNode);
                //case eControlType.Dialog:
                //    return new ExcelControlDialog(drawings, drawNode);
                default:
                    throw new NotSupportedException("Editboxes and Dialogs controls are not supported in worksheets");
            }
            if(ctrl is ExcelControlWithText t)
            {
                t.Text = name;
            }
            ctrl.Name = name;
            return ctrl;
        }
    }
}
