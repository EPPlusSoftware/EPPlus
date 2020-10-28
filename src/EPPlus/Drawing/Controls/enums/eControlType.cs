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
namespace OfficeOpenXml.Drawing.Controls
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
    public enum eControlType
    {
        Button,
        CheckBox,
        DropDown,
        GroupBox,
        Label,
        ListBox,
        RadioButton,
        ScrollBar,
        Spin,
        EditBox,
        Dialog
    }

}
