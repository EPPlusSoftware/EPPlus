﻿/*************************************************************************************************
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
    internal class IndexRelationWithSubRelations : IndexRelation
    {
        public IndexRelationWithSubRelations(RichDataIndexStore store, IndexEndpoint from, RichDataEntities subrelationEntity, IndexType indexType)
            : base(store, from, IndexEndpoint.None, indexType, IndexRelationType.SubRelations)
        {
            SubRelations = new List<IndexRelation>();
        }

        public RichDataEntities SubRelationEntity { get; }

        public virtual List<IndexRelation> SubRelations { get; }

        public bool HasSubrelations => SubRelations != null && SubRelations.Any();
    }
}
