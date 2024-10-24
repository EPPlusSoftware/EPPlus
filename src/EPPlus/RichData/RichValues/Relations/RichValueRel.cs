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
using OfficeOpenXml.Metadata;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.IndexRelations.EventArguments;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.Relations
{
    internal class RichValueRel : IndexEndpoint
    {
        public RichValueRel(ExcelWorkbook workbook, ZipPackagePart part) : base(workbook.IndexStore, RichDataEntities.RichValueRel)
        {
            _part = part;
            _workbook = workbook;
        }

        private readonly ZipPackagePart _part;
        private readonly ExcelWorkbook _workbook;

        public string RelationId { get; set; }
        public string Type { get; set; }

        public Uri TargetUri { get; set; }

        internal void WriteXml(StreamWriter sw)
        {
            sw.Write($"<rel r:id=\"{RelationId}\" />");
        }

        public override void DeleteMe(RelationDeletions relDeletions = null)
        {
            base.DeleteMe(relDeletions);
        }

        private void DeleteRelatedUri()
        {
            if (!_part.RelationshipExists(RelationId)) return;
            var rel = _part.GetRelationship(RelationId);
            if(rel.RelationshipType == ExcelPackage.schemaImage)
            {
                var pictureStore = _workbook._package.PictureStore;
                pictureStore.RemoveReference(TargetUri);
            }
            else
            {
                _part.DeleteRelationship(RelationId);
            }
        }

        public override void OnConnectedEntityDeleted(ConnectedEntityDeletedEventArgs e)
        {
            if (Deleted) return;
            base.OnConnectedEntityDeleted(e);
            if (e.DeletedEntity.EntityType == RichDataEntities.RichValue)
            {
                var rels = GetIncomingRelations();
                if(rels.Count() <= 1)
                {
                    DeleteRelatedUri();
                    // this was the last rich value connected to this relation
                    DeleteMe(e.RelationDeletions);
                }
            }
        }
    }
}
