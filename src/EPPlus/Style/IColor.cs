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
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Style
{
    //<xsd:complexType name = "CT_Color" >
    //    <xsd:attribute name = "auto" type="xsd:boolean" use="optional"/>
    //    <xsd:attribute name = "indexed" type="xsd:unsignedInt" use="optional"/>
    //    <xsd:attribute name = "rgb" type="ST_UnsignedIntHex" use="optional"/>
    //    <xsd:attribute name = "theme" type="xsd:unsignedInt" use="optional"/>
    //    <xsd:attribute name = "tint" type="xsd:double" use="optional" default="0.0"/>
    //</xsd:complexType>

    interface IColor
    {
        bool Auto { get;  }  
        int Indexed { get; set; }
        string Rgb { get; }
        eThemeSchemeColor? Theme { get; }
        decimal Tint { get; set; }
        void SetColor(Color color);
        void SetColor(eThemeSchemeColor color);
        void SetColor(ExcelIndexedColor color);
        void SetAuto();
    }
}
