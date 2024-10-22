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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class IndexRelation : IdentityItem
    {
        public IndexRelation(RichDataIndexStore store, IndexEndpoint from, IndexEndpoint to, IndexType indexType, IndexRelationType relationType = IndexRelationType.Default)
            : base(store)
        {
            From = from;
            To = to;
            IndexType = indexType;
            RelationType = relationType;
        }

        public bool Deleted { get; set; }

        public IndexEndpoint From { get; set; }

        public IndexEndpoint To { get; set; }

        public IndexType IndexType { get; set; }

        public IndexRelationType RelationType { get; set; }

        public IndexRelationWithSubRelations Parent { get; set; }

        public IndexRelationWithSubRelations AsRelationWithSubRelations()
        {
            return this as IndexRelationWithSubRelations;
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            var otherEndpoint = obj as IndexRelation;
            if (otherEndpoint == null) return false;
            return otherEndpoint.Id == Id;
        }

        public override int GetHashCode()
        {
            return Id.GetHashCode();
        }

        public static bool operator ==(IndexRelation left, IndexRelation right)
        {
            if (ReferenceEquals(left, null))
            {
                return ReferenceEquals(right, null);
            }

            return left.Equals(right);
        }

        public static bool operator !=(IndexRelation left, IndexRelation right)
        {
            return !(left == right);
        }
    }
}
