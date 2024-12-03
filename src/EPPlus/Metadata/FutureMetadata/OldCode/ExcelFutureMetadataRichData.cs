///*************************************************************************************************
//  Required Notice: Copyright (C) EPPlus Software AB. 
//  This software is licensed under PolyForm Noncommercial License 1.0.0 
//  and may only be used for noncommercial purposes 
//  https://polyformproject.org/licenses/noncommercial/1.0.0/

//  A commercial license to use this software can be purchased at https://epplussoftware.com
// *************************************************************************************************
//  Date               Author                       Change
// *************************************************************************************************
//  07/25/2024         EPPlus Software AB       EPPlus 7
// *************************************************************************************************/
//using OfficeOpenXml.Constants;
//using OfficeOpenXml.Utils;
//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Xml;

//namespace OfficeOpenXml.Metadata.FutureMetadata
//{
//    internal class ExcelFutureMetadataRichDataX : ExcelFutureMetadataType
//    {
//        public ExcelFutureMetadataRichData(int index)
//        {
//            Index = index;
//        }
//        public ExcelFutureMetadataRichData(XmlReader xr)
//        {
//            var startDepth = xr.Depth;
//            while (xr.Read() && startDepth <= xr.Depth)
//            {
//                if (xr.IsElementWithName("rvb"))
//                {
//                    Index = int.Parse(xr.GetAttribute("i"));
//                }
//            }

//            if (xr.NodeType == XmlNodeType.EndElement) xr.Read();
//        }
//        public int Index { get; private set; }
//        public override FutureMetadataType Type => FutureMetadataType.RichData;
//        public override string Uri => ExtLstUris.RichValueDataUri;
//        internal override void WriteXml(StreamWriter sw)
//        {
//            sw.Write($"<xlrd:rvb i=\"{Index}\"/>");
//        }
//    }
//}
