/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/25/2024         EPPlus Software AB       EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Metadata.FutureMetadata
{
    internal abstract class FutureMetadataBase : IndexEndpoint
    {
        internal const string DYNAMIC_ARRAY_NAME = "XLDAPR";
        internal const string RICHDATA_NAME = "XLRICHVALUE";

        protected FutureMetadataBase(RichDataIndexStore store) : base(store, RichDataEntities.FutureMetadata)
        {
        }

        public int Index { get; set; }
        public string Name { get; set; }

        public abstract string Uri { get; set; }

        public abstract IndexedCollection<FutureMetadataBlock> Blocks { get; set; }

        public abstract void Save(StreamWriter sw);
    }
}
