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
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    internal class ExcelDrawingCustomGeometry : XmlHelper
    {
        ExcelDrawing _drawing;
        internal ExcelDrawingCustomGeometry(ExcelDrawing drawing, XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            _drawing = drawing;
            var pathNode = GetNode("a:pathLst");
            if (pathNode != null)
            {
                foreach (var cn in pathNode.ChildNodes)
                {
                    if (cn is XmlElement e)
                    {
                        DrawingPaths.Add(new DrawingPath(e, NameSpaceManager));
                    }
                }
            }
        }
        internal List<DrawingPath> DrawingPaths { get; } = new List<DrawingPath>();

    }
}