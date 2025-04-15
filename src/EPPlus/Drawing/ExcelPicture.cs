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
using System.Text;
using System.Xml;
using System.IO;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Packaging;
using System.Linq;
using System.Globalization;
using OfficeOpenXml.Utils.Image;
using OfficeOpenXml.Utils.FileUtils;

#if NETFULL
using System.Drawing.Imaging;
using System.Xml.Linq;
#endif
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An image object
    /// </summary>
    public sealed class ExcelPicture : ExcelDrawing, IPictureContainer
    {
        #region "Constructors"
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Uri hyperlink, ePictureType type, PictureLocation location = PictureLocation.Embed, DrawingsCollectionType DrawingsType = DrawingsCollectionType.Worksheet) :
            base(drawings, node, NamespacePrefixes[(int)DrawingsType] + ":pic/", NamespacePrefixes[(int)DrawingsType] + ":nvPicPr/" + NamespacePrefixes[(int)DrawingsType] + ":cNvPr", null, DrawingsType)
        {
            Init();
            if (DrawingsType == DrawingsCollectionType.Chart)
            {
                node.OwnerDocument.DocumentElement.SetAttribute("xmlns:cdr", ExcelPackage.schemaChartDrawing);
                node.OwnerDocument.DocumentElement.SetAttribute("xmlns:a", ExcelPackage.schemaDrawings);
            }
            LocationType = location;

            bool containsEmbed = (location & PictureLocation.Embed) == PictureLocation.Embed;
            string attribute = containsEmbed ? "embed" : "link";

            var picNode = CreateNode(NamespacePrefixes[_prefixIndex] + ":pic");
            picNode.InnerXml = PicStartXml(type, attribute);

            Hyperlink = hyperlink;
            Image = new ExcelImage(this);
            Image.Type = type;

            switch (DrawingsType)
            {
                case DrawingsCollectionType.Chart:
                    int x = (int)(_drawings._screenWidth * EMU_PER_PIXEL * (From.X));
                    int y = (int)(_drawings._screenHeight * EMU_PER_PIXEL * (From.Y));
                    int cx = (int)(_drawings._screenWidth * EMU_PER_PIXEL * (To.X - From.X));
                    int cy = (int)(_drawings._screenHeight * EMU_PER_PIXEL * (To.Y - From.Y));
                    XmlElement xFrmNode = GetXfrmNode(picNode);
                    if (xFrmNode.ChildNodes.Count == 0)
                    {
                        CreateNode(xFrmNode, "a:off");
                        CreateNode(xFrmNode, "a:ext");
                    }
                    var offNode = (XmlElement)xFrmNode.SelectSingleNode("a:off", NameSpaceManager);
                    offNode.SetAttribute("x", x.ToString());
                    offNode.SetAttribute("y", y.ToString());
                    var extNode = (XmlElement)xFrmNode.SelectSingleNode("a:ext", NameSpaceManager);
                    extNode.SetAttribute("cx", cx.ToString());
                    extNode.SetAttribute("cy", cy.ToString());
                    Position = new ExcelDrawingCoordinate(drawings.NameSpaceManager, offNode, GetPositionSize);
                    Size = new ExcelDrawingSize(drawings.NameSpaceManager, extNode, GetPositionSize);
                    break;
                case DrawingsCollectionType.Worksheet:
                default:
                    node.AppendChild(picNode.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings));
                    break;
            }
        }

        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, ExcelGroupShape shape = null, DrawingsCollectionType DrawingsType = DrawingsCollectionType.Worksheet) :
            base(drawings, node, shape == null ? NamespacePrefixes[(int)DrawingsType] + ":pic/" : "", NamespacePrefixes[(int)DrawingsType] + ":nvPicPr/" + NamespacePrefixes[(int)DrawingsType] + ":cNvPr", shape, DrawingsType)
        {
            Init();
            XmlNode picNode = node.SelectSingleNode($"{_topPath}{NamespacePrefixes[_prefixIndex]}:blipFill/a:blip", drawings.NameSpaceManager);

            if (picNode != null)
            {

                var embedAttr = picNode.Attributes["embed", ExcelPackage.schemaRelationships];
                if (embedAttr != null && string.IsNullOrEmpty(embedAttr.Value) == false)
                {
                    LocationType = LocationType | PictureLocation.Embed;
                    IPictureContainer container = this;
                    container.RelPic = drawings.Part.GetRelationship(embedAttr.Value);
                    container.UriPic = UriHelper.ResolvePartUri(drawings.UriDrawing, container.RelPic.TargetUri);

                    if (drawings.Part.Package.PartExists(container.UriPic))
                    {
                        Part = drawings.Part.Package.GetPart(container.UriPic);
                    }
                    else
                    {
                        Part = null;
                        return;
                    }
                    ContentType = Part.ContentType;

                    var ms = ((MemoryStream)Part.GetStream());
                    Image = new ExcelImage(this);

                    var type = PictureStore.GetPictureTypeByContentType(ContentType);
                    if (type == null)
                    {
                        type = ImageReader.GetPictureType(ms, false);
                    }
                    Image.Type = type.Value;
                    byte[] iby = ms.ToArray();
                    Image.ImageBytes = iby;
                    var ii = _drawings._package.PictureStore.LoadImage(iby, container.UriPic, Part);
                    var pd = (IPictureRelationDocument)_drawings;
                    if (pd.Hashes.ContainsKey(ii.Hash))
                    {
                        pd.Hashes[ii.Hash].RefCount++;
                    }
                    else
                    {
                        pd.Hashes.Add(ii.Hash, new HashInfo(container.RelPic.Id) { RefCount = 1 });
                    }
                    container.ImageHash = ii.Hash;
                }

                var linkAttr = picNode.Attributes["link", ExcelPackage.schemaRelationships];

                if (linkAttr != null && string.IsNullOrEmpty(linkAttr.Value) == false)
                {
                    LocationType = LocationType | PictureLocation.Link;
                    LinkedImageRel = drawings.Part.GetRelationship(linkAttr.Value);
                    IPictureContainer container = this;

                    if (container.RelPic == null && container.UriPic == null)
                    {
                        container.RelPic = LinkedImageRel;
                        Image = new ExcelImage(this);
                        FileInfo ImageFile = new FileInfo(LinkedImageRel.TargetUri.LocalPath);
                        LoadImageLinked(ImageFile);
                    }
                }
            }
        }
        private void Init()
        {
            _lockAspectRatioPath = $"{_topPath}{NamespacePrefixes[_prefixIndex]}:nvPicPr/{NamespacePrefixes[_prefixIndex]}:cNvPicPr/a:picLocks/@noChangeAspect";
            _preferRelativeResizePath = $"{_topPath}{NamespacePrefixes[_prefixIndex]}:nvPicPr/{NamespacePrefixes[_prefixIndex]}:cNvPicPr/@preferRelativeResize";
            _rotationPath = string.Format(_rotationPath, _topPath, NamespacePrefixes[_prefixIndex]);
            _horizontalFlipPath = string.Format(_horizontalFlipPath, _topPath, NamespacePrefixes[_prefixIndex]);
            _verticalFlipPath = string.Format(_verticalFlipPath, _topPath, NamespacePrefixes[_prefixIndex]);
        }

        internal string GetRelId()
        {
            XmlElement blip = (XmlElement)TopNode.SelectSingleNode($"{_topPath}xdr:blipFill/a:blip", NameSpaceManager);

            return blip.GetAttribute("embed", ExcelPackage.schemaRelationships);
        }

        internal void SetRelId(XmlNode node, ePictureType type, string relID, string attribute = "embed")
        {
            XmlElement blip = (XmlElement)node.SelectSingleNode($"{_topPath}{NamespacePrefixes[_prefixIndex]}:blipFill/a:blip", NameSpaceManager);
            XmlElement blipSvg = null;
            if (type == ePictureType.Svg)
            {
                blipSvg = (XmlElement)node.SelectSingleNode($"{_topPath}{NamespacePrefixes[_prefixIndex]}:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip", NameSpaceManager);
                blipSvg.SetAttribute(attribute, ExcelPackage.schemaRelationships, relID);
            }

            blip.SetAttribute(attribute, ExcelPackage.schemaRelationships, relID);
        }

        /// <summary>
        /// The type of drawing
        /// </summary>
        public override eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.Picture;
            }
        }
#if !NET35 && !NET40
        internal async Task LoadImageAsync(Stream stream, ePictureType type)
        {
            var img = new byte[stream.Length];
            stream.Seek(0, SeekOrigin.Begin);
            await stream.ReadAsync(img, 0, (int)stream.Length).ConfigureAwait(false);

            SaveImageToPackage(type, img);
        }
#endif
        internal void LoadImage(Stream stream, ePictureType type)
        {
            var img = new byte[stream.Length];
            stream.Seek(0, SeekOrigin.Begin);
            var r = stream.Read(img, 0, (int)stream.Length);

            SaveImageToPackage(type, img);
        }

        internal void LoadImageWithoutSavingToPackage(Stream stream, ePictureType type)
        {
            var img = new byte[stream.Length];
            stream.Seek(0, SeekOrigin.Begin);
            var r = stream.Read(img, 0, (int)stream.Length);

            if (type == ePictureType.Emz ||
               type == ePictureType.Wmz)
            {
                img = ImageReader.ExtractImage(img, out ePictureType? pt);
                if (pt == null)
                {
                    throw (new InvalidDataException($"Invalid image of type {type}"));
                }
                type = pt.Value;
            }

            using (var ms = RecyclableMemory.GetStream(img))
            {
                Image.Type = type;
                Image.ImageBytes = img;
                Image.Type = type;
                RecalcWidthHeight();
            }
        }

        internal void LoadImageLinked(FileInfo ImageFile)
        {
            var uri = new Uri($"file:///{string.Format(ImageFile.FullName, CultureInfo.InvariantCulture)}");
            var type = PictureStore.GetPictureType(ImageFile.Extension);
            if (ImageFile.Exists)
            {
                LoadImageWithoutSavingToPackage(new FileStream(ImageFile.FullName, FileMode.Open, FileAccess.Read), type);
            }

            ContentType = PictureStore.GetContentType(type.ToString());
            LinkedImageRel = _drawings.Part._rels.FirstOrDefault(x => x.TargetUri.OriginalString == uri.OriginalString);
            if (LinkedImageRel == null)
            {
                LinkedImageRel = _drawings.Part.CreateRelationship(uri, TargetMode.External, ExcelPackage.schemaRelationships + "/image");
            }
            SetRelId(TopNode, type, LinkedImageRel.Id, "link");
        }

        private void SaveImageToPackage(ePictureType type, byte[] img)
        {
            var package = _drawings.Worksheet._package.ZipPackage;
            if (type == ePictureType.Emz ||
               type == ePictureType.Wmz)
            {
                img = ImageReader.ExtractImage(img, out ePictureType? pt);
                if (pt == null)
                {
                    throw (new InvalidDataException($"Invalid image of type {type}"));
                }
                type = pt.Value;
            }

            ContentType = PictureStore.GetContentType(type.ToString());
            var newUri = GetNewUri(package, "/xl/media/image{0}." + type.ToString());
            var store = _drawings._package.PictureStore;
            var pc = _drawings as IPictureRelationDocument;
            var ii = store.AddImage(img, newUri, type);

            IPictureContainer container = this;
            container.UriPic = ii.Uri;
            string relId;

            Part = ii.Part;

            if (!pc.Hashes.ContainsKey(ii.Hash))
            {
                container.RelPic = _drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(_drawings.UriDrawing, ii.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                relId = container.RelPic.Id;
                pc.Hashes.Add(ii.Hash, new HashInfo(relId));
                AddNewPicture(img, relId);
            }
            else
            {
                relId = pc.Hashes[ii.Hash].RelId;
                var rel = _drawings.Part.GetRelationship(relId);
                container.UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            }

            pc.Hashes[ii.Hash].RefCount++;

            container.ImageHash = ii.Hash;

            using (var ms = RecyclableMemory.GetStream(img))
            {
                Image.ImageBytes = img;
                Image.Type = type;
                RecalcWidthHeight();
            }

            //Create relationship
            SetRelId(TopNode, type, relId);
            //TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relId;
            package.Flush();
        }

        internal void RecalcWidthHeight()
        {
            //Ensure image has a size.width and size.height based on 100% orignal image
            if (Image != null && Image.ImageBytes != null)
            {
                //Recalculates width/height and bounds to 100% width/height relative to original image size
                Image.Bounds = PictureStore.GetImageBounds(Image.ImageBytes, Image.Type.Value, _drawings._package);

                var width = 0d;
                var height = 0d;
                ImageUtil.CalculateDPI(Image, STANDARD_DPI, ref width, ref height);

                SetPosDefaults((float)width, (float)height);
                if (_collectionType == DrawingsCollectionType.Chart)
                {
                    SetSize((int)width, (int)height);
                }

                //Image.Bounds.Height = origHeight;
                //Image.Bounds.Width = origWidth;
            }
        }

        private void AddNewPicture(byte[] img, string relID)
        {
            var newPic = new ExcelDrawings.ImageCompare();
            newPic.image = img;
            newPic.relID = relID;
            //_drawings._pics.Add(newPic);
        }
        #endregion
        private void SetPosDefaults(float width, float height)
        {
            var prevEdit = EditAs;

            if (EditAs != eEditAs.Absolute && _collectionType == DrawingsCollectionType.Worksheet)
            {
                EditAs = eEditAs.OneCell;
            }

            SetPixelWidth(width);
            SetPixelHeight(height);

            _width = GetPixelWidth();
            _height = GetPixelHeight();

            EditAs = prevEdit;
        }

        internal void SetNewId(int newId)
        {
            SetXmlNodeInt(NamespacePrefixes[_prefixIndex] + ":pic/" + NamespacePrefixes[_prefixIndex] + ":nvPicPr/" + NamespacePrefixes[_prefixIndex] + ":cNvPr /@id", newId, null, false);
        }

        private string PicStartXml(ePictureType type, string attribute)
        {
            StringBuilder xml = new StringBuilder();

            xml.AppendFormat("<{0}:nvPicPr>", NamespacePrefixes[_prefixIndex]);
            xml.AppendFormat("<{1}:cNvPr id=\"{0}\" descr=\"\" />", Id, NamespacePrefixes[_prefixIndex]);
            xml.AppendFormat("<{0}:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></{0}:cNvPicPr></{0}:nvPicPr><{0}:blipFill>", NamespacePrefixes[_prefixIndex]);
            if (type == ePictureType.Svg)
            {
                xml.Append($"<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:{attribute}=\"\" cstate=\"print\"><a:extLst><a:ext uri=\"{{28A0092B-C50C-407E-A947-70E740481C1C}}\"><a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext><a:ext uri=\"{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:{attribute}=\"\"/></a:ext></a:extLst></a:blip>");
            }
            else
            {
                xml.Append($"<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:{attribute}=\"\" cstate=\"print\" />");
            }
            xml.AppendFormat("<a:stretch><a:fillRect /> </a:stretch> </{0}:blipFill> <{0}:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </{0}:spPr>", NamespacePrefixes[_prefixIndex]);

            return xml.ToString();
        }

        /// <summary>
        /// The image
        /// </summary>
        public ExcelImage Image
        {
            get;
        }
        internal string ContentType
        {
            get;
            set;
        }
        /// <summary>
        /// Set the size of the image in percent from the orginal size
        /// Note that resizing columns / rows after using this function will effect the size of the picture
        /// </summary>
        /// <param name="Percent">Percent</param>
        public override void SetSize(int Percent)
        {
            if (Image.ImageBytes == null || _drawings._collectionType == DrawingsCollectionType.Chart)
            {
                base.SetSize(Percent);
            }
            else
            {
                ImageUtil.CalculateDPI(Image, STANDARD_DPI, ref _width, ref _height);

                _width = (int)(_width * ((double)Percent / 100));
                _height = (int)(_height * ((double)Percent / 100));

                _doNotAdjust = true;
                SetPixelWidth(_width);
                SetPixelHeight(_height);
                _doNotAdjust = false;
            }
        }
        internal Packaging.ZipPackagePart Part;

        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(_drawings, NameSpaceManager, TopNode, $"{_topPath}{NamespacePrefixes[_prefixIndex]}:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(_drawings, NameSpaceManager, TopNode, $"{_topPath}{NamespacePrefixes[_prefixIndex]}:spPr/a:ln", SchemaNodeOrder);
                }
                return _border;
            }
        }
        ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    _effect = new ExcelDrawingEffectStyle(_drawings, NameSpaceManager, TopNode, $"{_topPath}{NamespacePrefixes[_prefixIndex]}:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        string _preferRelativeResizePath;
        /// <summary>
        /// Relative to original picture size
        /// </summary>
        public bool PreferRelativeResize
        {
            get
            {
                return GetXmlNodeBool(_preferRelativeResizePath);
            }
            set
            {
                SetXmlNodeBool(_preferRelativeResizePath, value);
            }
        }
        string _lockAspectRatioPath;
        /// <summary>
        /// Lock aspect ratio
        /// </summary>
        public bool LockAspectRatio
        {
            get
            {
                return GetXmlNodeBool(_lockAspectRatioPath);
            }
            set
            {
                SetXmlNodeBool(_lockAspectRatioPath, value);
            }
        }
        internal override void CellAnchorChanged()
        {
            base.CellAnchorChanged();
            if (_fill != null) _fill.SetTopNode(TopNode);
            if (_border != null) _border.TopNode = TopNode;
            if (_effect != null) _effect.TopNode = TopNode;
        }

        internal override void DeleteMe()
        {
            IPictureContainer container = this;
            _drawings._package.PictureStore.RemoveImage(container.ImageHash, this);
            base.DeleteMe();
        }
        /// <summary>
        /// Dispose the object
        /// </summary>
        public override void Dispose()
        {
            //base.Dispose();
            //Hyperlink = null;
            //_image.Dispose();
            //_image = null;            
        }
        void IPictureContainer.RemoveImage()
        {
            IPictureContainer container = this;
            var relDoc = (IPictureRelationDocument)_drawings;
            if (relDoc.Hashes.TryGetValue(container.ImageHash, out HashInfo hi))
            {
                if (hi.RefCount <= 1)
                {
                    relDoc.Package.PictureStore.RemoveImage(container.ImageHash, this);
                    if (container.RelPic != null)
                    {
                        relDoc.RelatedPart.DeleteRelationship(container.RelPic.Id);
                    }
                    relDoc.Hashes.Remove(container.ImageHash);
                }
                else
                {
                    hi.RefCount--;
                }
            }
        }

        void IPictureContainer.SetNewImage()
        {
            var relId = ((IPictureContainer)this).RelPic.Id;
            var picNode = (XmlElement)TopNode.SelectSingleNode($"{_topPath}{NamespacePrefixes[_prefixIndex]}:blipFill/a:blip", NameSpaceManager);
            picNode.SetAttribute("r:embed", relId);
            if (Image.Type == ePictureType.Svg)
            {
                var node = TopNode.SelectSingleNode($"{_topPath}{NamespacePrefixes[_prefixIndex]}:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip/@r:embed", NameSpaceManager);
                if (node == null)
                {
                    var newNode = TopNode.OwnerDocument.CreateElement("a", "extLst", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    picNode.AppendChild(newNode);
                    newNode.InnerXml = $"<a:ext uri=\"{{28A0092B-C50C-407E-A947-70E740481C1C}}\"><a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext><a:ext uri=\"{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"{relId}\"/></a:ext>";
                }
                else
                {
                    node.Value = relId;
                }
            }
            else
            {
                if (picNode.HasChildNodes)
                {
                    picNode.InnerXml = "";
                }
            }

            if (_drawings.Part.RelationshipExists(relId))
            {
                IPictureContainer thisPicture = this;
                var thePart = _drawings.Part.Package.GetPart(thisPicture.UriPic);
                if (Part != thePart)
                {
                    Part = thePart;
                }
            }
        }

        string IPictureContainer.ImageHash { get; set; }
        Uri IPictureContainer.UriPic { get; set; }
        Packaging.ZipPackageRelationship IPictureContainer.RelPic { get; set; }
        IPictureRelationDocument IPictureContainer.RelationDocument => _drawings;
        string _rotationPath = "{0}{1}:spPr/a:xfrm/@rot";
        /// <summary>
        /// Rotation angle in degrees. Positive angles are clockwise. Negative angles are counter-clockwise.
        /// Note that EPPlus will not size the image depending on the rotation, so some angles will reqire the <see cref="ExcelDrawing.From"/> and <see cref="ExcelDrawing.To"/> coordinates to be set accordingly.
        /// </summary>
        public double Rotation
        {
            get
            {
                return GetXmlNodeAngle(_rotationPath);
            }
            set
            {
                SetXmlNodeAngle(_rotationPath, value, "Rotation", -100000, 100000);
            }
        }
        string _horizontalFlipPath = "{0}{1}:spPr/a:xfrm/@flipH";
        /// <summary>
        /// If true, flips the picture horizontal about the center of its bounding box.
        /// </summary>
        public bool HorizontalFlip
        {
            get
            {
                return GetXmlNodeBool(_horizontalFlipPath);
            }
            set
            {
                SetXmlNodeBool(_horizontalFlipPath, value, false);
            }
        }
        string _verticalFlipPath = "{0}{1}:spPr/a:xfrm/@flipV";
        /// <summary>
        /// If true, flips the picture vertical about the center of its bounding box.
        /// </summary>
        public bool VerticalFlip
        {
            get
            {
                return GetXmlNodeBool(_verticalFlipPath);
            }
            set
            {
                SetXmlNodeBool(_verticalFlipPath, value, false);
            }
        }

        internal PictureLocation LocationType = PictureLocation.None;

        internal ZipPackageRelationship LinkedImageRel = null;
    }
}