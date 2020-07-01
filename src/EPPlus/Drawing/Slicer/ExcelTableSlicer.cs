/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/26/2020         EPPlus Software AB       EPPlus 5.3
 ******0*******************************************************************************************/
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    /*
      <xsd:complexType name="CT_Slicer">
       <xsd:sequence>
         <xsd:element name="extLst" type="x:CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
       </xsd:sequence>
       <xsd:attribute name="name" type="x:ST_Xstring" use="required"/>
       <xsd:attribute ref="xr10:uid" use="optional"/>
       <xsd:attribute name="cache" type="x:ST_Xstring" use="required"/>
       <xsd:attribute name="caption" type="x:ST_Xstring" use="optional"/>
       <xsd:attribute name="startItem" type="xsd:unsignedInt" use="optional" default="0"/>
       <xsd:attribute name="columnCount" type="xsd:unsignedInt" use="optional" default="1"/>
       <xsd:attribute name="showCaption" type="xsd:boolean" use="optional" default="true"/>
       <xsd:attribute name="level" type="xsd:unsignedInt" use="optional" default="0"/>
       <xsd:attribute name="style" type="x:ST_Xstring" use="optional"/>
       <xsd:attribute name="lockedPosition" type="xsd:boolean" use="optional" default="false"/>
       <xsd:attribute name="rowHeight" type="xsd:unsignedInt" use="required"/>
     </xsd:complexType>
     */
    public class ExcelTableSlicer : ExcelSlicer<ExcelTableSlicerCache>
    {
        internal ExcelTableSlicer(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null) : base(drawings, node, parent)
        {

        }
    }
}
