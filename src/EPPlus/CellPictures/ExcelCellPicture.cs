/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.RichData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;

namespace OfficeOpenXml.CellPictures
{
    /// <summary>
    /// Represents an in-cell picture
    /// </summary>
    internal class ExcelCellPicture : IPictureContainer
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public ExcelCellPicture()
        {
            
        }

        public ExcelImage Image
        {
            get;
        }

        internal const string LocalImageStructureType = "_localImage";

        public ExcelAddress CellAddress { get; set;  }

        /// <summary>
        /// Name of the image file
        /// </summary>
       public string ImagePath { get; set; }

        public IPictureRelationDocument RelationDocument => throw new NotImplementedException();

        string IPictureContainer.ImageHash { get; set; }
        Uri IPictureContainer.UriPic { get; set; }
        ZipPackageRelationship IPictureContainer.RelPic { get; set; }
        internal int CalcOrigin { get; set; }

        public void RemoveImage()
        {
            //IPictureContainer container = this;
            //var relDoc = (IPictureRelationDocument)_drawings;
            //if (relDoc.Hashes.TryGetValue(container.ImageHash, out HashInfo hi))
            //{
            //    if (hi.RefCount <= 1)
            //    {
            //        relDoc.Package.PictureStore.RemoveImage(container.ImageHash, this);
            //        relDoc.RelatedPart.DeleteRelationship(container.RelPic.Id);
            //        relDoc.Hashes.Remove(container.ImageHash);
            //    }
            //    else
            //    {
            //        hi.RefCount--;
            //    }
            //}
        }

        public void SetNewImage()
        {
            //var relId = ((IPictureContainer)this).RelPic.Id;
            //TopNode.SelectSingleNode($"{_topPath}xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relId;
            //if (Image.Type == ePictureType.Svg)
            //{
            //    TopNode.SelectSingleNode($"{_topPath}xdr:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip/@r:embed", NameSpaceManager).Value = relId;
            //}
        }
    }
}
