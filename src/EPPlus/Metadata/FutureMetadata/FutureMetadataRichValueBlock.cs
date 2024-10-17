using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataRichValueBlock : FutureMetadataBlock
    {
        public FutureMetadataRichValueBlock(RichDataIndexStore store)
            : base(store, RichDataEntities.FutureMetadataRichDataBlock)
        {
            
        }
        public FutureMetadataRichValueBlock(XmlReader xr, RichDataIndexStore store)
            : base(store, RichDataEntities.FutureMetadataRichDataBlock)
        {
            ReadXml(xr);
        }

        private int? _indexFromRead;

        public override void InitRelations(ExcelRichData richData)
        {
            if(_indexFromRead.HasValue)
            {
                base.InitRelations();
                var rel = richData.Values.CreateRelation(this, richData.Values[_indexFromRead.Value], IndexType.ZeroBasedPointer);
                RichDataId = rel.To.Id;
            }
           
        }


        private void ReadXml(XmlReader xr)
        {
            while(!xr.EOF)
            {
                if (xr.IsElementWithName("rvb"))
                {
                    _indexFromRead = int.Parse(xr.GetAttribute("i"));
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

        public override void Save(StreamWriter sw)
        {
            var val = GetFirstTargetByType<ExcelRichValue>();
            if(val != null)
            {
                var ix = val.CurrentIndex;
                sw.Write("<bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\">");
                sw.Write($"<xlrd:rvb i=\"{ix}\" />");
                sw.Write("</ext></extLst></bk>");
            }
        }

        public int RichDataId { get; set; }
    }
}
