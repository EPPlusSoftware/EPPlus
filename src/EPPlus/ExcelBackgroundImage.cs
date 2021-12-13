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
using System.Xml;
using System.Drawing;
using System.IO;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;
namespace OfficeOpenXml
{
    /// <summary>
    /// An image that fills the background of the worksheet.
    /// </summary>
    public class ExcelBackgroundImage : XmlHelper
    {
        ExcelWorksheet _workSheet;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="nsm"></param>
        /// <param name="topNode">The topnode of the worksheet</param>
        /// <param name="workSheet">Worksheet reference</param>
        internal  ExcelBackgroundImage(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet workSheet) :
            base(nsm, topNode)
        {
            _workSheet = workSheet;
        }
        
        const string BACKGROUNDPIC_PATH = "d:picture/@r:id";
        /// <summary>
        /// The background image of the worksheet. 
        /// The image will be saved internally as a jpg.
        /// </summary>
        public Image Image
        {
            get
            {
                string relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
                if (!string.IsNullOrEmpty(relID))
                {
                    var rel = _workSheet.Part.GetRelationship(relID);
                    var imagePart = _workSheet.Part.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                    return Image.FromStream(imagePart.GetStream());
                }
                return null;
            }
            set
            {
                DeletePrevImage();
                if (value == null)
                {
                    DeleteAllNode(BACKGROUNDPIC_PATH);
                }
                else
                {
#if (Core)
                    var img=ImageCompat.GetImageAsByteArray(value);
#else
                    ImageConverter ic = new ImageConverter();
                    byte[] img = (byte[])ic.ConvertTo(value, typeof(byte[]));
#endif
                    var ii = _workSheet.Workbook._package.PictureStore.AddImage(img);
                    var rel = _workSheet.Part.CreateRelationship(ii.Uri, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                    SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
                }
            }
        }
        /// <summary>
        /// Set the picture from an image file. 
        /// The image file will be saved as a blob, so make sure Excel supports the image format.
        /// </summary>
        /// <param name="PictureFile">The image file.</param>
        public void SetFromFile(FileInfo PictureFile)
        {
            DeletePrevImage();

            byte[] fileBytes;
            fileBytes = File.ReadAllBytes(PictureFile.FullName);

            string contentType = PictureStore.GetContentType(PictureFile.Extension);
            var imageURI = XmlHelper.GetNewUri(_workSheet._package.ZipPackage, "/xl/media/" + PictureFile.Name.Substring(0, PictureFile.Name.Length - PictureFile.Extension.Length) + "{0}" + PictureFile.Extension);

            var ii = _workSheet.Workbook._package.PictureStore.AddImage(fileBytes, imageURI, contentType);


            if (_workSheet.Part.Package.PartExists(imageURI) && ii.RefCount==1) //The file exists with another content, overwrite it.
            {
                //Remove the part if it exists
                _workSheet.Part.Package.DeletePart(imageURI);
            }

            var imagePart = _workSheet.Part.Package.CreatePart(imageURI, contentType, CompressionLevel.None);
            //Save the picture to package.

            var strm = imagePart.GetStream(FileMode.Create, FileAccess.Write);
            strm.Write(fileBytes, 0, fileBytes.Length);

            var rel = _workSheet.Part.CreateRelationship(imageURI, Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
        }
        private void DeletePrevImage()
        {
            var relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
            if (relID != "")
            {
#if (Core)
                var img=ImageCompat.GetImageAsByteArray(Image);
#else
                var ic = new ImageConverter();
                byte[] img = (byte[])ic.ConvertTo(Image, typeof(byte[]));
#endif
                var ii = _workSheet.Workbook._package.PictureStore.GetImageInfo(img);

                //Delete the relation
                _workSheet.Part.DeleteRelationship(relID);
                
                //Delete the image if there are no other references.
                if (ii != null && ii.RefCount == 1)
                {
                    if (_workSheet.Part.Package.PartExists(ii.Uri))
                    {
                        _workSheet.Part.Package.DeletePart(ii.Uri);
                    }
                }
                
            }
        }
    }
}



