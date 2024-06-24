using OfficeOpenXml.OLE_Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.OleObject
{
    internal class OleObjectDrawing : ExcelDrawing
    {

        //TODO:
        //SET UPT DRAWING
        //REF TO OLE OBJECT
        //REF TO VML

        internal OleObjectDrawing(ExcelOleObject oleObject, ExcelDrawings drawings, XmlNode node, string topPath, string nvPrPath, ExcelGroupShape parent = null) : base(drawings, node, topPath, nvPrPath, parent)
        {

        }

        public override eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.OleObject;
            }
        }
    }
}
