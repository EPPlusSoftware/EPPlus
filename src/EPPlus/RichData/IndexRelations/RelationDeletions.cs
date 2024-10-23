using System;
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
