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
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Packaging;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// A picture fill for a drawing
    /// </summary>
    public class ExcelDrawingBlipFill : ExcelDrawingFillBase, IPictureContainer
    {
        string[] _schemaNodeOrder;
        private readonly IPictureRelationDocument _pictureRelationDocument;
        internal ExcelDrawingBlipFill(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nsm, XmlNode topNode, string fillPath, string[] schemaNodeOrder, Action initXml) : base(nsm, topNode, fillPath, initXml)
        {
            _schemaNodeOrder = schemaNodeOrder;
            _pictureRelationDocument = pictureRelationDocument;
            GetXml();
        }
        Image _image;
        /// <summary>
        /// The picture used in the fill.
        /// </summary>
        public Image Image
        {
            get
            {
                return _image;
            }
            set
            {
                if (_image == value) return;
                _initXml?.Invoke();
                if (_image != null)
                {
                    var container = (IPictureContainer)this;
                    _pictureRelationDocument.Package.PictureStore.RemoveImage(container.ImageHash, this);
                    _pictureRelationDocument.RelatedPart.DeleteRelationship(container.RelPic.Id);
                    _pictureRelationDocument.Hashes.Remove(container.ImageHash);
                }
                if (value != null)
                {
                    _image = value;
                    try
                    {
                        string relId = PictureStore.SavePicture(value, this);

                        //Create relationship
                        _xml.SetXmlNodeString("a:blip/@r:embed", relId);
                    }
                    catch (Exception ex)
                    {
                        throw (new Exception("Can't save image - " + ex.Message, ex));
                    }
                }
            }
        }
        /// <summary>
        /// Image format
        /// If the picture is created from an Image this type is always Jpeg
        /// </summary>
        public ImageFormat ImageFormat { get; internal set; } = ImageFormat.Jpeg;
        /// <summary>
        /// The image should be stretched to fill the target.
        /// </summary>
        public bool Stretch { get; set; } = false;
        /// <summary>
        /// Offset in percentage from the edge of the shapes bounding box. This property only apply when Stretch is set to true.        
        /// <seealso cref="Stretch"/>
        /// </summary>
        public ExcelDrawingRectangle StretchOffset { get; private set; } = new ExcelDrawingRectangle(0);
        /// <summary>
        /// The portion of the image to be used for the fill.
        /// Offset values are in percentage from the borders of the image
        /// </summary>
        public ExcelDrawingRectangle SourceRectangle { get; private set; } = new ExcelDrawingRectangle(0);
        /// <summary>
        /// The image should be tiled to fill the available space
        /// </summary>
        public ExcelDrawingBlipFillTile Tile
        {
            get;
            private set;
        } = new ExcelDrawingBlipFillTile();
        /// <summary>
        /// The type of fill
        /// </summary>
        public override eFillStyle Style
        {
            get
            {
                return eFillStyle.BlipFill;
            }
        }
        ExcelDrawingBlipEffects _effects=null;
        /// <summary>
        /// Blip fill effects
        /// </summary>
        public ExcelDrawingBlipEffects Effects
        {
            get
            {
                if(_effects==null)
                {
                    _effects = new ExcelDrawingBlipEffects(_nsm, _topNode.SelectSingleNode("a:blip", _nsm));
                }
                return _effects;
            }
        }
        internal override string NodeName
        {
            get
            {
                return "a:blipFill";
            }
        }

        internal override void GetXml()
        {
            var relId = _xml.GetXmlNodeString("a:blip/@r:embed");
            if (!string.IsNullOrEmpty(relId))
            {
                _image = PictureStore.GetPicture(relId, this, out string contentType);
                ContentType = contentType;
            }
            SourceRectangle = new ExcelDrawingRectangle(_xml, "a:srcRect/", 0);
            Stretch = _xml.ExistsNode("a:stretch");
            if (Stretch)
            {
                StretchOffset = new ExcelDrawingRectangle(_xml, "a:stretch/a:fillRect/", 0);
            }

            Tile = new ExcelDrawingBlipFillTile(_xml);
        }

        internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
        {
            _initXml?.Invoke();
            if (_xml == null) InitXml(nsm, node.FirstChild, "");
            CheckTypeChange(NodeName);

            if (SourceRectangle.BottomOffset != 0) _xml.SetXmlNodePercentage("a:srcRect/@b", SourceRectangle.BottomOffset);
            if (SourceRectangle.TopOffset != 0) _xml.SetXmlNodePercentage("a:srcRect/@t", SourceRectangle.TopOffset);
            if (SourceRectangle.LeftOffset != 0) _xml.SetXmlNodePercentage("a:srcRect/@l", SourceRectangle.LeftOffset);
            if (SourceRectangle.RightOffset != 0) _xml.SetXmlNodePercentage("a:srcRect/@r", SourceRectangle.RightOffset);

            if (Tile.Alignment != null && Tile.FlipMode != null)
            {
                if (Tile.Alignment.HasValue) _xml.SetXmlNodeString("a:tile/@algn", Tile.Alignment.Value.TranslateString());
                if (Tile.FlipMode.HasValue) _xml.SetXmlNodeString("a:tile/@flip", Tile.FlipMode.Value.ToString().ToLower());
                _xml.SetXmlNodePercentage("a:tile/@sx", Tile.HorizontalRatio, false);
                _xml.SetXmlNodePercentage("a:tile/@sy", Tile.VerticalRatio, false);
                _xml.SetXmlNodeString("a:tile/@tx", (Tile.HorizontalOffset * ExcelDrawing.EMU_PER_PIXEL).ToString(CultureInfo.InvariantCulture));
                _xml.SetXmlNodeString("a:tile/@ty", (Tile.VerticalOffset * ExcelDrawing.EMU_PER_PIXEL).ToString(CultureInfo.InvariantCulture));
            }

            if (Stretch)
            {
                _xml.SetXmlNodePercentage("a:stretch/a:fillRect/@b", StretchOffset.BottomOffset);
                _xml.SetXmlNodePercentage("a:stretch/a:fillRect/@t", StretchOffset.TopOffset);
                _xml.SetXmlNodePercentage("a:stretch/a:fillRect/@l", StretchOffset.LeftOffset);
                _xml.SetXmlNodePercentage("a:stretch/a:fillRect/@r", StretchOffset.RightOffset);
            }
        }

        internal override void UpdateXml()
        {
            SetXml(_xml.NameSpaceManager, _xml.TopNode);
        }

        internal void AddImage(FileInfo file)
        {
            if (!file.Exists)
            {
                throw (new ArgumentException($"File {file.FullName} does not exist."));
            }
            ContentType = PictureStore.GetContentType(file.Extension);
            var image = Image.FromFile(file.FullName);
            AddImage(image);
        }
        internal void AddImage(Image image)
        {
            Image = image;
        }
        #region IPictureContainer

        string IPictureContainer.ImageHash
        {
            get;
            set;
        }
        Uri IPictureContainer.UriPic
        {
            get;
            set;
        }
        ZipPackageRelationship IPictureContainer.RelPic
        {
            get;
            set;
        }


        internal string ContentType
        {
            get;
            set;
        }

        IPictureRelationDocument IPictureContainer.RelationDocument { get => _pictureRelationDocument; }
        #endregion
    }
}
