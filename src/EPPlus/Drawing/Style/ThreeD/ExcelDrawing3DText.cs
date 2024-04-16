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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.ThreeD
{
    /// <summary>
    /// 3D Text settings
    /// </summary>
    public class ExcelDrawing3DText : ExcelDrawing3D
    {
        private readonly string _flatTextZCoordinatePath = "{0}/a:flatTx/@z";
        internal ExcelDrawing3DText(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, path, schemaNodeOrder)
        {
            _flatTextZCoordinatePath = string.Format(_flatTextZCoordinatePath, path);
        }

        /// <summary>
        /// The Z coordinate to be used when positioning the flat text within the 3D scene
        /// </summary>
        public double FlatTextZCoordinate
        {
            get
            {
                return GetXmlNodeEmuToPtNull(_flatTextZCoordinatePath) ?? 0;
            }
            set
            {
                InitXml(true);
                SetXmlNodeEmuToPt(_flatTextZCoordinatePath, value);
            }
        }
    }
}
