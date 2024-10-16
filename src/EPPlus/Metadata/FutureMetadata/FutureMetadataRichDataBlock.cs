using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataRichDataBlock : IndexEndpoint
    {
        public FutureMetadataRichDataBlock(ExcelRichData richData)
            : base(richData.IndexStore, RichDataEntities.FutureMetadataRichDataBlock)
        {
            
        }
        public FutureMetadataRichDataBlock(XmlReader xr, ExcelRichData richData)
            : base(richData.IndexStore, RichDataEntities.FutureMetadataRichDataBlock)
        {
            _richData = richData;
            ReadXml(xr);
        }

        private readonly ExcelRichData _richData;

        private void ReadXml(XmlReader xr)
        {
            while(!xr.EOF)
            {
                if (xr.IsElementWithName("rvb"))
                {
                    var ix = int.Parse(xr.GetAttribute("i"));
                    var rel = _richData.Values.CreateRelation(this, _richData.Values[ix], IndexType.ZeroBasedPointer);
                    RichDataId = rel.To.Id;
                }
                else if(xr.IsEndElementWithName("bk"))
                {
                    break;
                }
                else
                {
                    xr.Read();
                }
            }
            
        }

        public void Save(StreamWriter sw)
        {
            var ix = _richData.Values.GetIndexById(RichDataId);
            if(ix.HasValue)
            {
                sw.Write("<bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\">");
                sw.Write($"<xlrd:rvb i=\"{ix}\" />");
                sw.Write("</ext></extLst></bk>");
            }
        }

        public int RichDataId { get; set; }
    }
}
