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
using System.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Fill properties for drawing objects
    /// </summary>
    public class ExcelDrawingFill : ExcelDrawingFillBasic
    {
        private readonly IPictureRelationDocument _pictureRelationDocument;
        internal ExcelDrawingFill(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string fillPath, string[] schemaNodeOrder) :
            base(pictureRelationDocument.Package, nameSpaceManager, topNode, fillPath, schemaNodeOrder, false)
        {
            _pictureRelationDocument = pictureRelationDocument;
            if (_fillNode != null)
            {
                LoadFill();
            }
        }
        /// <summary>
        /// Load the fill from the xml
        /// </summary>
        /// <param name="nameSpaceManager">The xml namespace manager</param>
        internal protected override void LoadFill()
        {
            if (_fillTypeNode == null) _fillTypeNode = _fillNode.SelectSingleNode("a:pattFill", NameSpaceManager);
            if (_fillTypeNode == null) _fillTypeNode = _fillNode.SelectSingleNode("a:blipFill", NameSpaceManager);

            if (_fillTypeNode == null)
            {
                base.LoadFill();
                return;
            }

            switch (_fillTypeNode.LocalName)
            {
                case "pattFill":
                    _style = eFillStyle.PatternFill;
                    _patternFill = new ExcelDrawingPatternFill(NameSpaceManager, _fillTypeNode, "", SchemaNodeOrder);
                    break;
                case "blipFill":
                    _style = eFillStyle.BlipFill;

                    _blipFill = new ExcelDrawingBlipFill(_pictureRelationDocument, NameSpaceManager, _fillTypeNode, "", SchemaNodeOrder);
                    break;
                default:
                    base.LoadFill();
                    break;
            }
        }

        internal override void SetFillProperty()
        {
            if (_fillNode == null)
            {
                base.SetFillProperty();
            }

            _patternFill = null;
            _blipFill = null;

            switch (_fillTypeNode.LocalName)
            {
                case "pattFill":
                    _patternFill    = new ExcelDrawingPatternFill(NameSpaceManager, _fillTypeNode, "", SchemaNodeOrder);
                    _patternFill.PatternType = eFillPatternStyle.Pct5;

                    if (_patternFill.BackgroundColor.ColorType == eDrawingColorType.None)
                    {
                        _patternFill.BackgroundColor.SetSchemeColor(eSchemeColor.Background1);
                    }
                    _patternFill.ForegroundColor.SetSchemeColor(eSchemeColor.Text1);
                    break;
                case "blipFill":
                    _blipFill = new ExcelDrawingBlipFill(_pictureRelationDocument, NameSpaceManager, _fillTypeNode, "", SchemaNodeOrder);                    
                    break;
                default:
                    base.SetFillProperty();
                    break;
            }
        }

        internal override void BeforeSave()
        {
            if (_patternFill != null)
            {
                PatternFill.UpdateXml();
            }
            else if (_blipFill != null)
            {
                BlipFill.UpdateXml();
            }
            else
            {
                base.BeforeSave();
            }
        }

        private ExcelDrawingPatternFill _patternFill = null;
        /// <summary>
        /// Reference pattern fill properties
        /// This property is only accessable when Type is set to PatternFill
        /// </summary>
        public ExcelDrawingPatternFill PatternFill
        {
            get
            {
                return _patternFill;
            }
        }
        private ExcelDrawingBlipFill _blipFill = null;
        /// <summary>
        /// Reference gradient fill properties
        /// This property is only accessable when Type is set to BlipFill
        /// </summary>
        public ExcelDrawingBlipFill BlipFill
        {
            get
            {
                return _blipFill;
            }
        }


        /// <summary>
        /// Disposes the object
        /// </summary>
        public new void Dispose()
        {
            base.Dispose();
            _patternFill = null;
            _blipFill = null;
        }
    }
}
