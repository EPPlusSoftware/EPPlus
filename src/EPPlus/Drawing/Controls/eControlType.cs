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
