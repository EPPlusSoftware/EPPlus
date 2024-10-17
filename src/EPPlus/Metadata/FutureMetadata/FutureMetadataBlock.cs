using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal abstract class FutureMetadataBlock : IndexEndpoint
    {
        protected FutureMetadataBlock(RichDataIndexStore store, RichDataEntities entity) : base(store, entity)
        {
        }

        public string Uri { get; set; }

        public abstract void Save(StreamWriter sw);

        public virtual void InitRelations(ExcelRichData richData)
        {

        }
    }
}
