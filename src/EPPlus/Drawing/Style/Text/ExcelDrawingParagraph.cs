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
using OfficeOpenXml.Style;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelDrawingParagraph
    {
        internal ExcelDrawingParagraph(XmlNamespaceManager nameSpaceManager, object pn, string[] schemaNodeOrder, Action initXml)
        {
        }
        public ExcelTextFont DefaultRunProperties 
        { 
            get; 
        }
        public ExcelDrawingTextRunCollection TextRuns 
        { 
            get;  
        }

        /*
         
3038 <xsd:attribute name="marL" type="ST_TextMargin" use="optional"/>
3039 <xsd:attribute name="marR" type="ST_TextMargin" use="optional"/>
3040 <xsd:attribute name="lvl" type="ST_TextIndentLevelType" use="optional"/>
3041 <xsd:attribute name="indent" type="ST_TextIndent" use="optional"/>
3042 <xsd:attribute name="algn" type="ST_TextAlignType" use="optional"/>
3043 <xsd:attribute name="defTabSz" type="ST_Coordinate32" use="optional"/>
3044 <xsd:attribute name="rtl" type="xsd:boolean" use="optional"/>
3045 <xsd:attribute name="eaLnBrk" type="xsd:boolean" use="optional"/>
3046 <xsd:attribute name="fontAlgn" type="ST_TextFontAlignType" use="optional"/>
3047 <xsd:attribute name="latinLnBrk" type="xsd:boolean" use="optional"/>
3048 <xsd:attribute name="hangingPunct" type="xsd:boolean" use="optional"/>         
         */
    }
}