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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.RichData.WebImages
{
    internal class WebImagesSupportingRichData : IndexEndpoint
    {
       

        public WebImagesSupportingRichData(ExcelWorkbook wb, ZipPackagePart webImagesPart, XmlReader xr)
            : base(wb.IndexStore, RichDataEntities.WebImage)
        {
            _wb = wb;
            _part = webImagesPart;
            ReadXml(xr);
        }

        public WebImagesSupportingRichData(ExcelWorkbook wb, ZipPackagePart webImagesPart)
            : base(wb.IndexStore, RichDataEntities.WebImage)
        {
            _wb = wb;
            _part = webImagesPart;
        }

        private readonly ExcelWorkbook _wb;
        private readonly ExcelRichData _richData;
        private readonly ZipPackagePart _part;
        private ZipPackageRelationship _addressRel;
        private ZipPackageRelationship _moreImagesAddressRel;
        private ZipPackageRelationship _blipRel;

        private void ReadXml(XmlReader xr)
        {
            do
            {
                if(xr.IsElementWithName("address"))
                {
                    var addressRelationId = xr.GetAttribute("r:id");
                    _addressRel = _part.GetRelationship(addressRelationId);
                }
                else if(xr.IsElementWithName("moreImagesAddress"))
                {
                    var moreImagesAddressRelationId = xr.GetAttribute("r:id");
                    _moreImagesAddressRel = _part.GetRelationship(moreImagesAddressRelationId);
                }
                else if(xr.IsElementWithName("blip"))
                {
                    var blipRelationId = xr.GetAttribute("r:id");
                    var blipRel = _part.GetRelationship(blipRelationId);
                }
                else if(xr.IsEndElementWithName("webImageSrd"))
                {
                    break;
                }
            }
            while (xr.Read());
        }

        public Uri Address => _addressRel.TargetUri;

        public Uri MoreImagesAddress => _moreImagesAddressRel?.TargetUri;

        /// <summary>
        /// BLIP (Binary Large Image or Picture). Uri to the local picture in the worksheet
        /// </summary>
        public Uri Blip => _blipRel?.TargetUri;

        internal void WriteXml(StreamWriter sw)
        {
            if(_addressRel != null && !string.IsNullOrEmpty(_addressRel.Id))
            {
                sw.Write($"<address r:id=\"{_addressRel.Id}\" />");
            }
            if(_moreImagesAddressRel != null && !string.IsNullOrEmpty(_moreImagesAddressRel.Id))
            {
                sw.Write($"<moreImagesAddress r:id=\"{_moreImagesAddressRel.Id}\" />");
            }
            if(_blipRel != null && !string.IsNullOrEmpty(_blipRel.Id))
            {
                sw.Write($"<blip ri:id=\"{_blipRel.Id}\" />");
            }
        }

        private void DeleteRelatedUris()
        {
            if(_blipRel != null && _part.RelationshipExists(_blipRel.Id))
            {
                var pictureStore = _wb._package.PictureStore;
                pictureStore.RemoveReference(_blipRel.TargetUri);
                _part.DeleteRelationship(_blipRel.Id);
            }
            if(_addressRel != null && _part.RelationshipExists(_addressRel.Id))
            {
                _part.DeleteRelationship(_addressRel.Id);
            }
            if(_moreImagesAddressRel != null && _part.RelationshipExists(_moreImagesAddressRel.Id))
            {
                _part.DeleteRelationship(_moreImagesAddressRel.Id);
            }
        }


        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            base.OnConnectedEntityDeleted(e);
            if(e.DeletedEntity.EntityType == RichDataEntities.RichValue)
            {
                var rels = GetIncomingRelations();
                if(rels.Count(x => x.From.EntityType == RichDataEntities.RichValue) <= 1)
                {
                    DeleteRelatedUris();
                    DeleteMe(e.RelationDeletions);
                }
            }
        }
    }
}
