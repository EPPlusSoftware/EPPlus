/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Style.Dxf;
using System;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// A style element for a custom table style with band size
    /// </summary>
    public class ExcelBandedTableStyleElement : ExcelTableStyleElement        
    {
        internal ExcelBandedTableStyleElement(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, eTableStyleElement type) : 
            base(nameSpaceManager, topNode, styles, type)
        {
            if(topNode!=null)
            {
                _bandSize= GetXmlNodeInt("@size",1);
            }
        }
        int _bandSize = 1;
        /// <summary>
        /// Band size. Only applicable when <see cref="Type"/> is set to FirstRowStripe, FirstColumnStripe, SecondRowStripe or SecondColumnStripe
        /// </summary>
        public int BandSize
        {
            get
            {
                return _bandSize;
            }
            set
            {
                if(value < 1 && value > 9)
                {
                    throw new InvalidOperationException("BandSize must be between 1 and 9");
                }
                _bandSize = value;
            }
        }
        internal override void CreateNode()
        {
            base.CreateNode();
            if (_bandSize == 1)
            {
                DeleteNode("@size");
            }
            else
            {
                SetXmlNodeInt("@size", _bandSize);
            }
        }
    }
}
