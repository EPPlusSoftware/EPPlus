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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An image object
    /// </summary>
    public sealed class ExcelPicture : ExcelDrawing
    {
        #region "Constructors"
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Uri hyperlink) :
            base(drawings, node, "xdr:pic", "xdr:nvPicPr/xdr:cNvPr")
        {
            CreatePicNode(node);
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
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Image image, Uri hyperlink) :
            base(drawings, node, "xdr:pic", "xdr:nvPicPr/xdr:cNvPr")
        {
            CreatePicNode(node);

            var package = drawings.Worksheet._package.ZipPackage;
            //Get the picture if it exists or save it if not.
            _image = image;
            Hyperlink = hyperlink;
            string relID = PictureStore.SavePicture(image, this);

            //Create relationship
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
            SetPosDefaults(image);
            package.Flush();
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
                _image = Image.FromStream(ms);
            }
            SetPosDefaults(_image);

            //Create relationship
            TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relId;
            package.Flush();
        }

        private void CreatePicNode(XmlNode node)
        {
            var picNode = CreateNode("xdr:pic");
            picNode.InnerXml = PicStartXml();

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
        private void SetPosDefaults(Image image)
        {
            EditAs = eEditAs.OneCell;
            var width = image.Width / (image.HorizontalResolution / STANDARD_DPI);
            var height = image.Height / (image.VerticalResolution / STANDARD_DPI);
            SetPixelWidth(width);
            SetPixelHeight(height);
            _width = GetPixelWidth();
            _height = GetPixelHeight();
        }

        private string PicStartXml()
        {
            StringBuilder xml = new StringBuilder();

            xml.Append("<xdr:nvPicPr>");
            xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", _id);
            xml.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" /><a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");

            return xml.ToString();
        }

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