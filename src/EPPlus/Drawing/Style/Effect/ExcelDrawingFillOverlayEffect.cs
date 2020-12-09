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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{
    /// <summary>
    /// The fill overlay effect. 
    /// A fill overlay can be used to specify an additional fill for a drawing and blend the two together.
    /// </summary>
    public class ExcelDrawingFillOverlayEffect : ExcelDrawingEffectBase
    {
        private readonly IPictureRelationDocument _pictureRelationDocument;
        internal ExcelDrawingFillOverlayEffect(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path) : base(nameSpaceManager, topNode, schemaNodeOrder, path)
        {
            _pictureRelationDocument = pictureRelationDocument;
        }
        /// <summary>
        /// The fill to blend with
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get;
            private set;
        }
        /// <summary>
        /// How to blend the overlay
        /// Default is Over
        /// </summary>
        public eBlendMode Blend
        {
            get
            {
                return GetXmlNodeString(_path+"/@blend").ToEnum(eBlendMode.Over);
            }
            set
            {
                if(Fill==null)
                {
                    Create();
                }
                SetXmlNodeString(_path + "/@blend", value.ToString().ToLowerInvariant());
            }
        }
        /// <summary>
        /// Creates a fill overlay with BlendMode = Over
        /// </summary>
        public void Create()
        {
            if (Fill == null)
            {                
                Fill = new ExcelDrawingFill(_pictureRelationDocument, NameSpaceManager, TopNode, _path, SchemaNodeOrder);
                if(Fill._fillTypeNode==null)
                {
                    Fill.Style = eFillStyle.NoFill;
                }
            }

            if (!ExistNode($"{_path}/@blend"))
            {
                Blend = eBlendMode.Over;
            }
        }
        /// <summary>
        /// Removes any fill overlay
        /// </summary>
        public void Remove()
        {
            DeleteNode($"{_path}");
            Fill = null;
        }

    }
}