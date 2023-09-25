﻿using OfficeOpenXml.Core;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class WritingManager
    {
        private readonly Dictionary<string, int> _styleCache = new Dictionary<string, int>();
        private readonly EPPlusReadOnlyList<ExcelRangeBase> _ranges;

        EpplusCssWriter cssWriter;
        EpplusHtmlWriter htmlwriter;

        Stream stream;

        public WritingManager(EPPlusReadOnlyList<ExcelRangeBase> ranges, HtmlRangeExportSettings settings) 
        {

            cssWriter = new EpplusCssWriter(RecyclableMemory.GetStream(), ranges._list, settings, settings.Css, settings.Css.CssExclude);
            //htmlwriter = new EpplusHtmlWriter(


        }
    }
}
