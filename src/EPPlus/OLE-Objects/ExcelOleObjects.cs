using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;

namespace OfficeOpenXml.OLE_Objects
{
    public class ExcelOleObjects : IEnumerable<ExcelOleObject>, IDisposable
    { 
        internal List<ExcelOleObject> _oleObjects = new List<ExcelOleObject>();
        internal ExcelWorksheet _worksheet;

        public ExcelOleObjects(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
            XmlNode node = worksheet.WorksheetXml.SelectSingleNode("/d:worksheet/d:oleObjects", worksheet.NameSpaceManager);
            if (node != null && worksheet != null)
            {
                AddOleObjects(node);
            }
        }

        private void AddOleObjects(XmlNode node)
        {
            XmlNodeList list = node.SelectNodes("mc:AlternateContent", _worksheet.NameSpaceManager);

            foreach (XmlNode n in list)
            {
                ExcelOleObject oo;
                switch (n.LocalName)
                {
                    case "AlternateContent":
                        var ole = n.SelectSingleNode("mc:Choice/d:oleObject", _worksheet.NameSpaceManager);
                        oo = ExcelOleObject.GetOleObject(_worksheet, _worksheet.NameSpaceManager, this, ole);
                        break;
                    default:
                        oo = null;
                        break;
                }
                if (oo != null)
                {
                    _oleObjects.Add(oo);
                }
            }
        }

        public ExcelOleObject this[int PositionID]
        {
            get
            {
                return (_oleObjects[PositionID]);
            }
        }

        internal ExcelOleObject GetOleObjectByShapeId(int shapeId)
        {
            return _oleObjects.FirstOrDefault(x => x.shapeId == shapeId);
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        public IEnumerator<ExcelOleObject> GetEnumerator()
        {
            return _oleObjects.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _oleObjects.GetEnumerator();
        }
    }
}
