using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style
{

    /// <summary>
    /// Used to define line-spacing in pargraphs/shapes
    /// Default is single line spacing
    /// </summary>
    public class ExcelLineSpacing : XmlHelper
    {
        Action _initXml;
        string _path, _subPath;
        internal ExcelLineSpacing(XmlNamespaceManager nsm, XmlNode topNode,string path, string[] schemaNodeOrder, Action initXml, eDrawingTextLineSpacing defLst) : base(nsm, topNode)
        {
            _path = path;
            _initXml = initXml;
            SchemaNodeOrder = schemaNodeOrder;
            var node = GetNode(path);
            if(node==null || node.HasChildNodes==false)
            {
                LineSpacingType = defLst;
            }
            else if (node.ChildNodes[0].LocalName== "spcPts")
            {
                _subPath = "a:spcPts";
                _lineSpacingType = eDrawingTextLineSpacing.Multiple;
            }
            else
            {
                _subPath = "a:spcPct";
                _lineSpacingType = eDrawingTextLineSpacing.Exactly;
            }
        }

        /// <summary>
        /// Value for the line spacing. 
        /// If the value is <see cref="eDrawingTextLineSpacing.Exactly"/> the value is in points.
        /// Otherwise the value is in percent. Default is 100 percent.
        /// </summary>
        public double Value 
        {
            get
            {
;               var v = GetXmlNodeDoubleNull($"{_path}{_subPath}/@val") ?? 100;
                if (_lineSpacingType == eDrawingTextLineSpacing.Exactly)
                {
                    return v / 100;
                }
                else
                {
                    return v / 1000;
                }
            }
                set
            {
                SetXmlNodeInt($"{_path}{_subPath}/@val", GetXmlValue(value));               
            }
        }

        private int GetXmlValue(double value)
        {
            if(_lineSpacingType==eDrawingTextLineSpacing.Exactly)
            {
                return (int)(value * 100);
            }
            else
            {
                return (int)(value * 1000);
            }
        }

        eDrawingTextLineSpacing _lineSpacingType;
        /// <summary>
        /// If setting Exactly or Multiple it is recommended to use
        /// SetExactly or SetMultiple functions.
        /// Otherwise they are set to default values 13.2 or 3
        /// </summary>
        public eDrawingTextLineSpacing LineSpacingType
        {
            get
            {
                return _lineSpacingType;
            }
            set
            {
                if (value == LineSpacingType) return;

                switch (_lineSpacingType)
                {
                    case eDrawingTextLineSpacing.Exactly:
                        SetExactly(13.2);
                        break;
                    case eDrawingTextLineSpacing.Multiple:
                        SetMultiple(300);
                        break;
                    case eDrawingTextLineSpacing.Double:
                        SetMultiple(200);
                        break;
                    case eDrawingTextLineSpacing.Single:
                        SetMultiple(100);
                        break;
                    case eDrawingTextLineSpacing.OneAndAHalf:
                        SetMultiple(150);
                        break;
                }
            }
        }

        /// <summary>
        /// Sets line spacing to Exactly inPoints
        /// </summary>
        /// <param name="inPoints"></param>
        public void SetExactly(double inPoints)
        {
            if(LineSpacingType != eDrawingTextLineSpacing.Exactly)
            {
                DeleteNode($"{_path}{_subPath}");
                _subPath = "a:spcPts";
            }
            _lineSpacingType = eDrawingTextLineSpacing.Exactly;
            Value = inPoints;
        }
        /// <summary>
        /// Sets line spacing to multiple of percent
        /// </summary>
        /// <param name="percent"></param>
        public void SetMultiple(double percent)
        {
            if (_lineSpacingType == eDrawingTextLineSpacing.Exactly)
            {
                DeleteNode($"{_path}{_subPath}");
                _subPath = "a:spcPct";
            }
            _lineSpacingType = eDrawingTextLineSpacing.Multiple;
            Value = percent;
        }

        ///// <summary>
        ///// If setting Exactly or Multiple it is recommended to use
        ///// SetExactly or SetMultiple functions.
        ///// Otherwise they are set to default values 13.2 or 3
        ///// </summary>
        ///// <param name="lineSpacingType"></param>
        //public void SetSpacingType(eDrawingTextLineSpacing lineSpacingType)
        //{
        //    LineSpacingType = lineSpacingType;
        //    if(lineSpacingType == eDrawingTextLineSpacing.Exactly)
        //    {
        //        SetExactly(13.2);
        //    }
        //    else if(lineSpacingType == eDrawingTextLineSpacing.Multiple)
        //    {
        //        SetMultiple(3);
        //    }
        //}
    }
}
