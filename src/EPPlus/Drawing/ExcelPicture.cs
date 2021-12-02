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
using System.Globalization;
using System.Text;
using System.Xml;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Packaging;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
#if Core
using SkiaSharp;
#endif
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An image object
    /// </summary>
    public sealed class ExcelPicture : ExcelDrawing
    {
#region "Constructors"
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Uri hyperlink, ePictureType type) :
            base(drawings, node, "xdr:pic", "xdr:nvPicPr/xdr:cNvPr")
        {
            CreatePicNode(node,type);
            Hyperlink = hyperlink;
        }

        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, ExcelGroupShape shape = null) :
            base(drawings, node, "xdr:pic", "xdr:nvPicPr/xdr:cNvPr", shape)
        {
            XmlNode picNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);
            if (picNode != null && picNode.Attributes["embed", ExcelPackage.schemaRelationships] != null)
            {
                IPictureContainer container = this;
                container.RelPic = drawings.Part.GetRelationship(picNode.Attributes["embed", ExcelPackage.schemaRelationships].Value);
                container.UriPic = UriHelper.ResolvePartUri(drawings.UriDrawing, container.RelPic.TargetUri);

                var extension = Path.GetExtension(container.UriPic.OriginalString);
                ContentType = PictureStore.GetContentType(extension);
                if (drawings.Part.Package.PartExists(container.UriPic))
                {
                    Part = drawings.Part.Package.GetPart(container.UriPic);
                }
                else
                {
                    Part = null;
                    _image = null;
                    return;
                }
#if (Core)
                try
                {
                    _image = Image.FromStream(Part.GetStream());
                }
                catch
                {
                    if(extension.ToLower()==".emf" || extension.ToLower() == ".wmf") //Not supported in linux environments, so we ignore them and set image to null.
                    {
                        _image = null;
                        return;
                    }
                    else
                    {
                        throw;
                    }
                }
                byte[] iby = ImageCompat.GetImageAsByteArray(_image);
#else
                _image = Image.FromStream(Part.GetStream());
                ImageConverter ic =new ImageConverter();
                var iby=(byte[])ic.ConvertTo(_image, typeof(byte[]));
#endif
                var ii = _drawings._package.PictureStore.LoadImage(iby, container.UriPic, Part);
                container.ImageHash = ii.Hash;
            }
        }
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Image image, Uri hyperlink, ePictureType type) :
            base(drawings, node, "xdr:pic", "xdr:nvPicPr/xdr:cNvPr")
        {
            CreatePicNode(node, type);

            var package = drawings.Worksheet._package.ZipPackage;
            //Get the picture if it exists or save it if not.
            _image = image;
            Hyperlink = hyperlink;
            string relID = PictureStore.SavePicture(image, this);

            SetRelId(node, type, relID);
            var width = image.Width / (image.HorizontalResolution / STANDARD_DPI);
            var height = image.Height / (image.VerticalResolution / STANDARD_DPI);
            SetPosDefaults(width, height);
            package.Flush();
        }

        private void SetRelId(XmlNode node, ePictureType type, string relID)
        {
            //Create relationship
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
            if (type == ePictureType.Svg)
            {
                node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip/@r:embed", NameSpaceManager).Value = relID;
            }
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
            stream.Read(img, 0, (int)stream.Length);

            SaveImageToPackage(type, img);
        }
        private void SaveImageToPackage(ePictureType type, byte[] img)
        {
            var package = _drawings.Worksheet._package.ZipPackage;
            ContentType = PictureStore.GetContentType(type.ToString());
            IPictureContainer container = this;
            container.UriPic = GetNewUri(package, "/xl/media/image{0}." + type.ToString());
            var store = _drawings._package.PictureStore;
            var pc = _drawings as IPictureRelationDocument;            
            var ii = store.AddImage(img, container.UriPic, ContentType);
            string relId;
            if (!pc.Hashes.ContainsKey(ii.Hash))
            {
                Part = ii.Part;
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
            container.ImageHash = ii.Hash;
            using (var ms = RecyclableMemory.GetStream(img))
            {
#if (Core)                
                
                double width=0, height=0;
                try
                {
                    if (type == ePictureType.Bmp ||
                       type == ePictureType.Jpg ||
                       type == ePictureType.Gif ||
                       type == ePictureType.Png ||
                       type == ePictureType.Tif)
                    {
                        var isImg = SixLabors.ImageSharp.Image.Load(ms);
                        var scale = GetImageDpi(isImg.Metadata.ResolutionUnits);
                        width = (float)(isImg.Width / (isImg.Metadata.HorizontalResolution * scale / STANDARD_DPI));
                        height = (float)(isImg.Height / (isImg.Metadata.VerticalResolution * scale / STANDARD_DPI));
                    }
                    else
                    {
                        TryGetImageBounds(type, ms, ref width, ref height);
                    }
                }
                catch
                {
                    TryGetImageBounds(type, ms, ref width, ref height);
                }
                SetPosDefaults((float)width, (float)height);

                float GetImageDpi(SixLabors.ImageSharp.Metadata.PixelResolutionUnit unit)
                {
                    switch (unit)
                    {
                        case SixLabors.ImageSharp.Metadata.PixelResolutionUnit.PixelsPerMeter:
                            return 1 / 39.36F;
                        case SixLabors.ImageSharp.Metadata.PixelResolutionUnit.PixelsPerCentimeter:
                            return 100 / 39.36F;
                        default:                    
                            return 1;

                    }
                }
#else
                _image = Image.FromStream(ms);
                var width = _image.Width / (_image.HorizontalResolution / STANDARD_DPI);
                var height = _image.Height / (_image.VerticalResolution / STANDARD_DPI);

                SetPosDefaults(width, height);
#endif
            }

            //Create relationship
            SetRelId(TopNode, type, relId);
            //TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relId;
            package.Flush();
        }

        private bool TryGetImageBounds(ePictureType pictureType, MemoryStream ms, ref double width, ref double height)
        {
            ms.Seek(0,SeekOrigin.Begin);
            if(pictureType==ePictureType.Emf && IsEmf(ms, ref width, ref height)) 
            {
                return true;
            }
            if (pictureType == ePictureType.Wmf && IsWmf(ms, ref width, ref height))
            {
                return true;
            }
            else if (pictureType==ePictureType.Svg && IsSvg(ms, ref width, ref height))
            {
                return true;
            }
            return false;
        }
        private bool IsEmf(MemoryStream ms, ref double width, ref double height)
        {
            var br = new BinaryReader(ms);
            if (br.PeekChar() == 1)
            {
                var type = br.ReadInt32();
                var length = br.ReadInt32();
                var bounds = new int[4];
                bounds[0] = br.ReadInt32();
                bounds[1] = br.ReadInt32();
                bounds[2] = br.ReadInt32();
                bounds[3] = br.ReadInt32();
                var frame = new int[4];
                frame[0] = br.ReadInt32();
                frame[1] = br.ReadInt32();
                frame[2] = br.ReadInt32();
                frame[3] = br.ReadInt32();

                var signatureBytes = br.ReadBytes(4);
                var signature = Encoding.ASCII.GetString(signatureBytes);
                if (signature.Trim() == "EMF")
                {
                    width = bounds[2] + 2;
                    height = bounds[3] + 2;
                    return true;
                }
            }
            return false;
        }

        private bool IsWmf(MemoryStream ms, ref double width, ref double height)
        {
            var br = new BinaryReader(ms);
            //if (br.PeekChar() == 0xD7)
            //{
                var key = br.ReadUInt32(); //0x9AC6CDD7
                var HWmf = br.ReadInt16();
                var bounds=new ushort[4];
                bounds[0]= br.ReadUInt16();
                bounds[1] = br.ReadUInt16();
                bounds[2] = br.ReadUInt16();
                bounds[3] = br.ReadUInt16();

                var inch = br.ReadInt16();
            //}
            return false;
        }

        private bool IsSvg(MemoryStream ms, ref double width, ref double height)
        {
            try
            {
                using (var reader = new XmlTextReader(ms))
                {
                    while (reader.Read())
                    {
                        if (reader.LocalName == "svg" && reader.NodeType == XmlNodeType.Element)
                        {
                            var w=reader.GetAttribute("width");
                            var h=reader.GetAttribute("height");
                            var vb = reader.GetAttribute("viewBox");
                            reader.Close();
                            if (w==null || h==null)
                            {
                                if (vb==null)
                                {
                                    return false;
                                }
                                var bounds = vb.Split(new char[]{ ' ',',' },StringSplitOptions.RemoveEmptyEntries);
                                if (bounds.Length < 4)
                                {
                                    return false;
                                }
                                if(string.IsNullOrEmpty(w))
                                {
                                    w = bounds[2];
                                }
                                if (string.IsNullOrEmpty(h))
                                {
                                    h = bounds[3];
                                }
                            }
                            width = GetSvgUnit(w);
                            if (double.IsNaN(width)) return false;
                            height = GetSvgUnit(h);
                            if (double.IsNaN(height)) return false;
                            return true;
                        }
                    }
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private double GetSvgUnit(string v)
        {
            var factor = 1D;
            if (v.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            {
                v = v.Substring(0, v.Length - 2);
            }
            else if(v.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                factor = 1.25;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("pc", StringComparison.OrdinalIgnoreCase))
            {
                factor = 15;
                v = v.Substring(0, v.Length - 2);
            }
            else if(v.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
            {
                factor = 3.543307;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            {
                factor = 35.43307;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            {
                factor = 90;
                v = v.Substring(0, v.Length - 2);
            }
            if(double.TryParse(v, out double value))
            {
                return value*factor;
            }
            return double.NaN;
        }
        private void CreatePicNode(XmlNode node, ePictureType type)
        {
            var picNode = CreateNode("xdr:pic");
            picNode.InnerXml = PicStartXml(type);

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings), picNode);
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
            EditAs = eEditAs.OneCell;
            SetPixelWidth(width);
            SetPixelHeight(height);
            _width = GetPixelWidth();
            _height = GetPixelHeight();
        }

        private string PicStartXml(ePictureType type)
        {
            StringBuilder xml = new StringBuilder();

            xml.Append("<xdr:nvPicPr>");
            xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", _id);
            xml.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>");
            if(type==ePictureType.Svg)
            {
                xml.Append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\"><a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext><a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"\"/></a:ext></a:extLst></a:blip>");
            }
            else
            {
                xml.Append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" />");
            }
            xml.Append("<a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");

            return xml.ToString();
        }

#if Core
        SKBitmap _imageSkia=null;
#endif
        Image _image = null;
        /// <summary>
        /// The Image
        /// </summary>
        public Image Image
        {
            get
            {
                return _image;
            }
            set
            {
                if (value != null)
                {
                    _image = value;
                    try
                    {
                        string relID = PictureStore.SavePicture(value, this);

                        //Create relationship
                        TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
                    }
                    catch (Exception ex)
                    {
                        throw (new Exception("Can't save image - " + ex.Message, ex));
                    }
                }
            }
        }
        ImageFormat _imageFormat = ImageFormat.Jpeg;
        /// <summary>
        /// Image format
        /// If the picture is created from an Image this type is always Jpeg
        /// </summary>
        public ImageFormat ImageFormat
        {
            get
            {
                return _imageFormat;
            }
            internal set
            {
                _imageFormat = value;
            }
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
            if (Image == null)
            {
                base.SetSize(Percent);
            }
            else
            {
                _width = Image.Width / (Image.HorizontalResolution / STANDARD_DPI);
                _height = Image.Height / (Image.VerticalResolution / STANDARD_DPI);

                _width = (int)(_width * ((double)Percent / 100));
                _height = (int)(_height * ((double)Percent / 100));

                _doNotAdjust = true;
                SetPixelWidth(_width);
                SetPixelHeight(_height);
                _doNotAdjust = false;
            }
        }
        internal Packaging.ZipPackagePart Part;

        internal new string Id
        {
            get { return Name; }
        }
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
                    _fill = new ExcelDrawingFill(_drawings, NameSpaceManager, TopNode, "xdr:pic/xdr:spPr", SchemaNodeOrder);
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
                    _border = new ExcelDrawingBorder(_drawings, NameSpaceManager, TopNode, "xdr:pic/xdr:spPr/a:ln", SchemaNodeOrder);
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
                    _effect = new ExcelDrawingEffectStyle(_drawings, NameSpaceManager, TopNode, "xdr:pic/xdr:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        const string _preferRelativeResizePath = "xdr:pic/xdr:nvPicPr/xdr:cNvPicPr/@preferRelativeResize";
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
        const string _lockAspectRatioPath = "xdr:pic/xdr:nvPicPr/xdr:cNvPicPr/a:picLocks/@noChangeAspect";
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
    }
}