using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using OfficeOpenXml.Drawing.OleObject;

namespace OfficeOpenXml
{
    internal class OleObjectsCollectionInternal : XmlHelper, IEnumerable<OleObjectInternal>
    {
        private List<OleObjectInternal> _list = new List<OleObjectInternal>();

        internal OleObjectsCollectionInternal(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            var nodes = GetNodes("d:oleObjects/mc:AlternateContent/mc:Choice/d:oleObject");
            foreach (XmlNode node in nodes)
            {
                _list.Add(new OleObjectInternal(NameSpaceManager, node));
            }
        }

        internal OleObjectInternal GetOleObjectByShapeId(int shapeId)
        {
            return _list.FirstOrDefault(x => x.ShapeId == shapeId);
        }

        public IEnumerator<OleObjectInternal> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}
