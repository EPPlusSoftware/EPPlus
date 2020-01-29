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

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// The border style of a drawing in a differential formatting record
    /// </summary>
    public class ExcelDxfBorderBase : DxfStyleBase<ExcelDxfBorderBase>
    {
        internal ExcelDxfBorderBase(ExcelStyles styles)
            : base(styles)
        {
            Left=new ExcelDxfBorderItem(_styles);
            Right = new ExcelDxfBorderItem(_styles);
            Top = new ExcelDxfBorderItem(_styles);
            Bottom = new ExcelDxfBorderItem(_styles);
        }
        /// <summary>
        /// Left border style
        /// </summary>
        public ExcelDxfBorderItem Left
        {
            get;
            internal set;
        }
        /// <summary>
        /// Right border style
        /// </summary>
        public ExcelDxfBorderItem Right
        {
            get;
            internal set;
        }
        /// <summary>
        /// Top border style
        /// </summary>
        public ExcelDxfBorderItem Top
        {
            get;
            internal set;
        }
        /// <summary>
        /// Bottom border style
        /// </summary>
        public ExcelDxfBorderItem Bottom
        {
            get;
            internal set;
        }

        /// <summary>
        /// The Id
        /// </summary>
        protected internal override string Id
        {
            get
            {
                return Top.Id + Bottom.Id + Left.Id + Right.Id/* + Diagonal.Id + GetAsString(DiagonalUp) + GetAsString(DiagonalDown)*/;
            }
        }

        /// <summary>
        /// Creates the the xml node
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The X Path</param>
        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            Left.CreateNodes(helper, path + "/d:left");
            Right.CreateNodes(helper, path + "/d:right");
            Top.CreateNodes(helper, path + "/d:top");
            Bottom.CreateNodes(helper, path + "/d:bottom");
        }
        /// <summary>
        /// If the object has a value
        /// </summary>
        protected internal override bool HasValue
        {
            get 
            {
                return Left.HasValue ||
                    Right.HasValue ||
                    Top.HasValue ||
                    Bottom.HasValue;
            }
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        protected internal override ExcelDxfBorderBase Clone()
        {
            return new ExcelDxfBorderBase(_styles) { Bottom = Bottom.Clone(), Top=Top.Clone(), Left=Left.Clone(), Right=Right.Clone() };
        }
    }
}
