using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Metadata
{
    internal class MdxMetadata : IndexEndpoint
    {
        private readonly string _xml;

        public MdxMetadata(RichDataIndexStore store) : base(store, RichDataEntities.MdxMetadata)
        {
        }

        public MdxMetadata(XmlReader reader, RichDataIndexStore store) : base(store, RichDataEntities.MdxMetadata)
        { 
            _xml = reader.ReadOuterXml();
        }


    }
}
