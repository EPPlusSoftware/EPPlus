﻿using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.RichData.TestClasses
{
    internal class FutureMetadataBlockTest : IndexEndpoint
    {
        public FutureMetadataBlockTest(RichDataIndexStore store) : base(store, RichDataEntities.FutureMetadataRichDataBlock)
        {
        }
    }
}
