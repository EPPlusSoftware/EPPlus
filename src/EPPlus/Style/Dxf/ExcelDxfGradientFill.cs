/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/29/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// Represents a gradient fill used for differential style formatting.
    /// </summary>
    public class ExcelDxfGradientFill : DxfStyleBase
    {
        internal ExcelDxfGradientFill(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
            : base(styles, callback)
        {
            Colors = new ExcelDxfGradientFillColorCollection(styles, callback);
        }

        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get
            {
                return Colors.HasValue || Degree.HasValue || Left.HasValue || Right.HasValue || Top.HasValue || Bottom.HasValue || GradientType.HasValue;
            }
        }
        internal override string Id 
        {
            get
            {
                return Colors.Id + "|" + GetAsString(Degree) + "|" + GetAsString(Left) + "|" + GetAsString(Right) + "|" + GetAsString(Top) + "|" + GetAsString(Bottom) + "|" + GetAsString(GradientType);
            }
        } 
        internal static string GetEmptyId()
        {
            return $"{ExcelDxfColor.GetEmptyId()}||||||";
        }
        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {
            Degree = null;
            Left = null;
            Right = null;
            Top = null;
            Bottom = null;
            Colors.Clear();
        }
        /// <summary>
        /// A collection of colors and percents for the gradient fill
        /// </summary>
        public ExcelDxfGradientFillColorCollection Colors 
        { 
            get;
            private set;
        }
        internal override DxfStyleBase Clone()
        {
            return new ExcelDxfGradientFill(_styles, _callback)
            {
                Colors = (ExcelDxfGradientFillColorCollection)Colors.Clone(),
                Degree = Degree,
                Left = Left,
                Right = Right,
                Top = Top,
                Bottom = Bottom
            };
        }
        eDxfGradientFillType? _gradientType;
        /// <summary>
        /// Type of gradient fill
        /// </summary>
        public eDxfGradientFillType? GradientType 
        { 
            get
            {
                return _gradientType;
            }
            set
            {
                _gradientType = value;
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientType, value);
            }
        }
        double? _degree;

        /// <summary>
        /// Angle of the linear gradient
        /// </summary>
        public double? Degree
        {
            get
            {
                return _degree;
            }
            set
            {
                _degree = value;
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientDegree, value);
            }
        }
        double? _left;

        /// <summary>
        /// The left position of the inner rectangle (color 1). 
        /// </summary>
        public double? Left
        {
            get
            {
                return _left;
            }
            set
            {
                _left = value;
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientLeft, value);
            }
        }

        double? _right;
        /// <summary>
        /// The right position of the inner rectangle (color 1). 
        /// </summary>
        public double? Right
        {
            get
            {
                return _right;
            }
            set
            {
                _right = value;
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientRight, value);
            }
        }

        double? _top;
        /// <summary>
        /// The top position of the inner rectangle (color 1). 
        /// </summary>
        public double? Top
        {
            get
            {
                return _top;
            }
            set
            {
                _top = value;
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientTop, value);
            }
        }
        double? _bottom;
        /// <summary>
        /// The bottom position of the inner rectangle (color 1). 
        /// </summary>
        public double? Bottom
        {
            get
            {
                return _bottom;
            }
            set
            {
                _bottom = value;
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientBottom, value);
            }
        }

        internal override void CreateNodes(XmlHelper helper, string path)
        {
            var gradNode = helper.CreateNode(path + "/d:gradientFill");
            var gradHelper=XmlHelperFactory.Create(helper.NameSpaceManager, gradNode);
            SetValueEnum(gradHelper, "@type", GradientType);
            SetValue(gradHelper, "@degree", Degree);
            SetValue(gradHelper, "@left", Left);
            SetValue(gradHelper, "@right", Right);
            SetValue(gradHelper, "@top", Top);
            SetValue(gradHelper, "@bottom", Bottom);

            foreach (var c in Colors)
            {
                c.CreateNodes(gradHelper, "");
            }
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientType, _gradientType);
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientDegree, _degree);
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientTop, _top);
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientBottom, _bottom);
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientLeft, _left);
                _callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientRight, _right);
                foreach (var c in Colors)
                {
                    c.SetStyle();
                }
            }
        }

        internal override void SetValuesFromXml(XmlHelper helper)
        {
            GradientType = helper.GetXmlNodeString("d:fill/d:gradientFill/@type").ToEnum<eDxfGradientFillType>();
            Degree = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@degree");
            Left = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@left");
            Right = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@right");
            Top = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@top");
            Bottom = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@bottom");
            foreach (XmlNode node in helper.GetNodes("d:fill/d:gradientFill/d:stop"))
            {
                var stopHelper = XmlHelperFactory.Create(_styles.NameSpaceManager, node);
                var c = Colors.Add(stopHelper.GetXmlNodeDouble("@position") * 100);
                c.Color = GetColor(stopHelper, "d:color", c.Position==0 ? eStyleClass.FillGradientColor1 : eStyleClass.FillGradientColor2);
            }
        }
    }
}