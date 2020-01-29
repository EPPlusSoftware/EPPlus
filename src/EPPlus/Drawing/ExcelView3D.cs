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
using System.Text;
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// 3D settings
    /// </summary>
    public sealed class ExcelView3D : XmlHelper
    {
       internal ExcelView3D(XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           SchemaNodeOrder = new string[] { "rotX", "hPercent", "rotY", "depthPercent","rAngAx", "perspective"};
       }
       const string perspectivePath = "c:perspective/@val";
       /// <summary>
       /// Degree of perspective 
       /// </summary>
       public decimal Perspective
       {
           get
           {
               return GetXmlNodeInt(perspectivePath);
           }
           set
           {
               SetXmlNodeString(perspectivePath, value.ToString(CultureInfo.InvariantCulture));
           }
       }
       const string rotXPath = "c:rotX/@val";
       /// <summary>
       /// Rotation X-axis
       /// </summary>
       public decimal RotX
       {
           get
           {
               return GetXmlNodeDecimal(rotXPath);
           }
           set
           {
               CreateNode(rotXPath);
               SetXmlNodeString(rotXPath, value.ToString(CultureInfo.InvariantCulture));
           }
       }
       const string rotYPath = "c:rotY/@val";
       /// <summary>
       /// Rotation Y-axis
       /// </summary>
       public decimal RotY
       {
           get
           {
               return GetXmlNodeDecimal(rotYPath);
           }
           set
           {
               CreateNode(rotYPath);
               SetXmlNodeString(rotYPath, value.ToString(CultureInfo.InvariantCulture));
           }
       }
       const string rAngAxPath = "c:rAngAx/@val";
       /// <summary>
       /// Right Angle Axes
       /// </summary>
       public bool RightAngleAxes
       {
           get
           {
               return GetXmlNodeBool(rAngAxPath);
           }
           set
           {
               SetXmlNodeBool(rAngAxPath, value);
           }
       }
       const string depthPercentPath = "c:depthPercent/@val";
       /// <summary>
       /// Depth % of base
       /// </summary>
       public int DepthPercent
       {
           get
           {
               return GetXmlNodeInt(depthPercentPath);
           }
           set
           {
               if (value < 0 || value > 2000)
               {
                   throw(new ArgumentOutOfRangeException("Value must be between 0 and 2000"));
               }
               SetXmlNodeString(depthPercentPath, value.ToString());
           }
       }
       const string heightPercentPath = "c:hPercent/@val";
       /// <summary>
       /// Height % of base
       /// </summary>
       public int HeightPercent
       {
           get
           {
               return GetXmlNodeInt(heightPercentPath);
           }
           set
           {
               if (value < 5 || value > 500)
               {
                   throw (new ArgumentOutOfRangeException("Value must be between 5 and 500"));
               }
               SetXmlNodeString(heightPercentPath, value.ToString());
           }
       }
    }
}
