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
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// A collection of colors and their positions used for a gradiant fill.
    /// </summary>
    public class ExcelDrawingGradientFillColorList : IEnumerable<ExcelDrawingGradientFillColor>
    {
        List<ExcelDrawingGradientFillColor> _lst = new List<ExcelDrawingGradientFillColor>();
        private XmlNamespaceManager _nsm;
        private XmlNode _topNode;
        private XmlNode _gsLst=null;
        private string _path;
        private string[] _schemaNodeOrder;
        internal ExcelDrawingGradientFillColorList(XmlNamespaceManager nsm, XmlNode topNode, string path, string[] schemaNodeOrder)
        {
            _nsm = nsm;
            _topNode = topNode;
            _path = path;
            _schemaNodeOrder = schemaNodeOrder;
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index in the collection</param>
        /// <returns>The color</returns>
        public ExcelDrawingGradientFillColor this[int index]
        {
            get
            {
                return (_lst[index]);
            }
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
        /// <summary>
        /// Gets the first occurance with the color with the specified position
        /// </summary>
        /// <param name="position">The position in percentage</param>
        /// <returns>The color</returns>
        public ExcelDrawingGradientFillColor this[double position]
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
        /// <param name="color">The Color</param>
        public void AddRgb(double position, Color color)
        {
            var gs = GetGradientFillColor(position);
            gs.Color.SetRgbColor(color);
            _lst.Add(gs);
        }
        /// <summary>
        /// Adds a RGB percentage color at the specified position
        /// </summary>
        /// <param name="position">The position</param>
        /// <param name="redPercentage">The percentage of red</param>
        /// <param name="greenPercentage">The percentage of green</param>
        /// <param name="bluePercentage">The percentage of blue</param>
        public void AddRgbPercentage(double position, double redPercentage, double greenPercentage, double bluePercentage)
        {
            var gs = GetGradientFillColor(position);
            gs.Color.SetRgbPercentageColor(redPercentage, greenPercentage, bluePercentage);
            _lst.Add(gs);
        }
        /// <summary>
        /// Adds a theme color at the specified position
        /// </summary>
        /// <param name="position">The position</param>
        /// <param name="color">The theme color</param>
        public void AddScheme(double position, eSchemeColor color)
        {
            var gs = GetGradientFillColor(position);
            gs.Color.SetSchemeColor(color);
            _lst.Add(gs);
        }
        /// <summary>
        /// Adds a system color at the specified position
        /// </summary>
        /// <param name="position">The position</param>
        /// <param name="color">The system color</param>
        public void AddSystem(double position, eSystemColor color)
        {
            var gs = GetGradientFillColor(position);
            gs.Color.SetSystemColor(color);
            _lst.Add(gs);
        }
        /// <summary>
        /// Adds a HSL color at the specified position
        /// </summary>
        /// <param name="position">The position</param>
        /// <param name="hue">The hue part. Ranges from 0-360</param>
        /// <param name="saturation">The saturation part. Percentage</param>
        /// <param name="luminance">The luminance part. Percentage</param>
        public void AddHsl(double position, double hue, double saturation, double luminance)
        {            
            var gs = GetGradientFillColor(position);
            gs.Color.SetHslColor(hue, saturation, luminance);
            _lst.Add(gs);
        }
        /// <summary>
        /// Adds a HSL color at the specified position
        /// </summary>
        /// <param name="position">The position</param>
        /// <param name="color">The preset color</param>
        public void AddPreset(double position, ePresetColor color)
        {
            var gs = GetGradientFillColor(position);
            gs.Color.SetPresetColor(color);
            _lst.Add(gs);
        }

        private ExcelDrawingGradientFillColor GetGradientFillColor(double position)
        {
            if (position < 0 || position > 100)
            {
                throw (new ArgumentOutOfRangeException("Position must be between 0 and 100"));
            }
            XmlNode node = null;
            for (var i = 0; i < _lst.Count; i++)
            {
                if (_lst[i].Position > position)
                {
                    node = AddGs(position, _lst[i].TopNode);
                }
            }
            node = AddGs(position, null);

            var tc = new ExcelDrawingGradientFillColor()
            {
                Position = position,
                Color = new ExcelDrawingColorManager(_nsm, node, "", _schemaNodeOrder),
                TopNode = node
            };
            return tc;
        }

        private XmlElement AddGs(double position, XmlNode node)
        {
            if(_gsLst==null)
            {
                var xml = XmlHelperFactory.Create(_nsm, _topNode);
                _gsLst=xml.CreateNode(_path);
            }
            var gs = _gsLst.OwnerDocument.CreateElement("a", "gs", ExcelPackage.schemaDrawings);
            if (node == null)
            {
                _gsLst.AppendChild(gs);
            }
            else
            {
                _gsLst.InsertBefore(gs, node);
            }
            gs.SetAttribute("pos", (position * 1000).ToString());
            return gs;
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelDrawingGradientFillColor> GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        internal void Add(double position, XmlNode node)
        {
            _lst.Add(new ExcelDrawingGradientFillColor()
            {
                Position = position,
                Color = new ExcelDrawingColorManager(_nsm, node, "", _schemaNodeOrder),
                TopNode = node
            });
        }
    }
}
