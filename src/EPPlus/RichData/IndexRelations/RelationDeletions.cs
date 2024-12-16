using System;
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
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.IndexRelations
{
    internal class RelationDeletions
    {
        public RelationDeletions(RichDataIndexStore store)
        {
            _store = store;
        }

        private readonly RichDataIndexStore _store;
        private readonly Stack<IndexRelation> _deletions = new Stack<IndexRelation>();

        public void EnqueueDelete(IndexRelation relationToDelete)
        {
            if(!relationToDelete.Deleted)
            {
                relationToDelete.Deleted = true;
                _deletions.Push(relationToDelete);
            }
        }

        public void DeleteAll()
        {
            while(_deletions.Count > 0)
            {
                var rel = _deletions.Pop();
                _store.DeleteRelation(rel);
            }
        }
    }
}
