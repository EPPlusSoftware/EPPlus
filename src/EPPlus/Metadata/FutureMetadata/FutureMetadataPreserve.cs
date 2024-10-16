using OfficeOpenXml.RichData;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal class FutureMetadataPreserve : ExcelFutureMetadata
    {
        public FutureMetadataPreserve(string name, int count, ExcelRichData richData)
            : base(richData.IndexStore)
        {
            _name = name;
            _count = count;
        }

        public override string Uri { get; set; }

        private string _innerXml;
        private readonly string _name;
        private readonly int _count;

        public void ReadXml(XmlReader xr)
        {
            _innerXml = xr.ReadInnerXml();
        }

        protected override void Save(StreamWriter sw)
        {
            sw.Write($"<futureMetadata name=\"{_name}\" count=\"{_count}\">");
            sw.Write(_innerXml);
            sw.Write("</futureMetadata>");
        }
    }
}
