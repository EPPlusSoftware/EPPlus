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
    internal class FutureMetadataPreserve : FutureMetadataBase
    {
        public FutureMetadataPreserve(string name, int count, RichDataIndexStore store)
            : base(store)
        {
            _name = name;
            _count = count;
        }

        public FutureMetadataPreserve(XmlReader xr, RichDataIndexStore store)
            : base(store)
        {
            if(xr.IsElementWithName("futureMetadata"))
            {
                _name = xr.GetAttribute("name");
                _count = int.Parse(xr.GetAttribute("count"));
            }
        }

        public override string Uri { get; set; }
        public override IndexedCollection<FutureMetadataBlock> Blocks 
        { 
            get
            {
                return default;    
            }
            set
            {
            }
        }

        private string _innerXml;
        private readonly string _name;
        private readonly int _count;

        public void ReadXml(XmlReader xr)
        {
            _innerXml = xr.ReadInnerXml();
        }

        public override void Save(StreamWriter sw)
        {
            sw.Write($"<futureMetadata name=\"{_name}\" count=\"{_count}\">");
            sw.Write(_innerXml);
            sw.Write("</futureMetadata>");
        }
    }
}
