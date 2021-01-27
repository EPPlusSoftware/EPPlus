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
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Utils.Extensions;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfGradientFill : DxfStyleBase
    {
        internal ExcelDxfGradientFill(ExcelStyles styles)
            : base(styles)
        {
            Colors = new ExcelDxfGradientFillColorList(styles);
        }

        public override bool HasValue
        {
            get
            {
                return Colors.HasValue || Degree.HasValue || Left.HasValue || Right.HasValue || Top.HasValue || Bottom.HasValue || GradientType.HasValue;
            }
        }

        protected internal override string Id 
        {
            get
            {
                return Colors.Id + "|" + GetAsString(Degree) + "|" + GetAsString(Left) + "|" + GetAsString(Right) + "|" + GetAsString(Top) + "|" + GetAsString(Bottom) + "|" + GetAsString(GradientType);
            }
        } 

        public override void Clear()
        {
            Degree = null;
            Left = null;
            Right = null;
            Top = null;
            Bottom = null;
            Colors.Clear();
        }
        public ExcelDxfGradientFillColorList Colors 
        { 
            get;
            private set;
        }
        protected internal override DxfStyleBase Clone()
        {
            return new ExcelDxfGradientFill(_styles)
            {
                Colors = (ExcelDxfGradientFillColorList)Colors.Clone(),
                Degree = Degree,
                Left = Left,
                Right = Right,
                Top = Top,
                Bottom = Bottom
            };
        }
        public eDxfGradientFillType? GradientType { get; set; }
        public double? Degree { get; set; }
        public double? Left { get; set; }
        public double? Right { get; set; }
        public double? Top { get; set; }
        public double? Bottom { get; set; }
        protected internal override void CreateNodes(XmlHelper helper, string path)
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
        protected internal override void SetValuesFromXml(XmlHelper helper)
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
                var c = Colors.Add(stopHelper.GetXmlNodeDouble("@position"));
                c.Color = GetColor(stopHelper, "d:color");
            }
        }
    }
    /// <summary>
    /// A collection of colors and their positions used for a gradiant fill.
    /// </summary>
    public class ExcelDxfGradientFillColorList : DxfStyleBase, IEnumerable<ExcelDxfGradientFillColor>
    {
        List<ExcelDxfGradientFillColor> _lst = new List<ExcelDxfGradientFillColor>();
        public ExcelDxfGradientFillColorList(ExcelStyles styles) : base(styles)
        {
            _styles = styles;
        }
        public IEnumerator<ExcelDxfGradientFillColor> GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _lst.GetEnumerator();
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index in the collection</param>
        /// <returns>The color</returns>
        public ExcelDxfGradientFillColor this[int index]
        {
            get
            {
                return (_lst[index]);
            }
        }
        /// <summary>
        /// Gets the first occurance with the color with the specified position
        /// </summary>
        /// <param name="position">The position in percentage</param>
        /// <returns>The color</returns>
        public ExcelDxfGradientFillColor this[double position]
        {
            get
            {
                return (_lst.Find(i => i.Position == position));
            }
        }
        /// <summary>
        /// Adds a RGB color at the specified position
        /// </summary>
        /// <param name="position">The position</param>
        /// <returns>The gradient color position object</returns>
        public ExcelDxfGradientFillColor Add(double position)
        {
            var color = new ExcelDxfGradientFillColor(_styles, position);
            _lst.Add(color);
            return color;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _lst.Count;
            }
        }

        protected internal override string Id => throw new System.NotImplementedException();

        public override bool HasValue
        {
            get
            {
                return _lst.Count > 0;
            }
        }

        public void RemoveAt(int index)
        {
            _lst.RemoveAt(index);
        }
        public void RemoveAt(double position)
        {
            var item = _lst.Find(i => i.Position == position);
            if(item!=null)
            {
                _lst.Remove(item);
            }
        }
        public void Remove(ExcelDxfGradientFillColor item)
        {
            _lst.Remove(item);
        }
       public override void Clear()
        {
            _lst.Clear();
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            if(_lst.Count>0)
            {
                foreach(var c in _lst)
                {
                    c.CreateNodes(helper, path);
                }
            }
        }

        protected internal override DxfStyleBase Clone()
        {
            var ret = new ExcelDxfGradientFillColorList(_styles);
            foreach (var c in _lst)
            {
                ret._lst.Add((ExcelDxfGradientFillColor)c.Clone());
            }
            return ret;
        }
    }
    public class ExcelDxfGradientFillColor : DxfStyleBase
    {
        internal ExcelDxfGradientFillColor(ExcelStyles styles, double position)
            : base(styles)
        {
            Position = position;
            Color = new ExcelDxfColor(styles);
        }
        public double Position { get; }
        public ExcelDxfColor Color { get; internal set; }

        public override bool HasValue
        {
            get
            {
                return Color.HasValue;
            }
        }

        protected internal override string Id => Position.ToString() + "|" + Color.Id;

        public override void Clear()
        {
            Color.Clear();
        }

        protected internal override DxfStyleBase Clone()
        {
            return new ExcelDxfGradientFillColor(_styles, Position)
            {
                Color = (ExcelDxfColor)Color.Clone()
            };
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            var node = helper.CreateNode(path + "d:stop", false, true);
            var stopHelper = XmlHelperFactory.Create(helper.NameSpaceManager, node);
            SetValue(stopHelper, "@position", Position);
            SetValueColor(stopHelper, "d:color", Color);
        }
    }
}