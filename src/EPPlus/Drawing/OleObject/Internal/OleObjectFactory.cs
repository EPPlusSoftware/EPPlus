/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing.OleObject
{
    internal static class OleObjectFactory
    {
        internal static ExcelDrawing GetOleObject(ExcelDrawings drawings, XmlElement drawNode, OleObjectInternal oleObject, ExcelGroupShape parent)
        {
            XmlNode node;
            if (parent == null)
            {
                node = drawNode.ParentNode;
            }
            else
            {
                node = drawNode;
            }
            return new ExcelOleObject(drawings, node, oleObject, parent);
        }

        internal static ExcelOleObject CreateOleObject(ExcelDrawings drawings, XmlElement drawNode, string name, string olePath, ExcelOleObjectParameters parameters)
        {
            return new ExcelOleObject(drawings, drawNode, name, olePath, parameters);
        }
        internal static ExcelOleObject CreateOleObject(ExcelDrawings drawings, XmlElement drawNode, string name, FileInfo oleInfo, ExcelOleObjectParameters parameters)
        {
            return new ExcelOleObject(drawings, drawNode, name, oleInfo, parameters);
        }
        internal static ExcelOleObject CreateOleObject(ExcelDrawings drawings, XmlElement drawNode, string name, Stream oleStream, ExcelOleObjectParameters parameters)
        {
            return new ExcelOleObject(drawings, drawNode, name, oleStream, parameters);
        }
    }
}
