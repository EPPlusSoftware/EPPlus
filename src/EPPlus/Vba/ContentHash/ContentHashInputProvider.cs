/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/05/2022         EPPlus Software AB       EPPlus 6.1
 *************************************************************************************************/
using OfficeOpenXml.VBA;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Vba.ContentHash
{
    internal abstract class ContentHashInputProvider
    {
        public ContentHashInputProvider(ExcelVbaProject project)
        {
            _project = project;
            _hashEncoding = System.Text.Encoding.GetEncoding(Project.CodePage);
        }

        private readonly ExcelVbaProject _project;
        private readonly Encoding _hashEncoding; 

        protected ExcelVbaProject Project => _project;
        protected Encoding HashEncoding => _hashEncoding;

        public void CreateHashInput(MemoryStream ms)
        {
            if(ms == null)
            {
                ms = RecyclableMemory.GetStream();
            }
            CreateHashInputInternal(ms);
        }

        protected abstract void CreateHashInputInternal(MemoryStream s);

        public static void GetContentNormalizedDataHashInput(ExcelVbaProject project, MemoryStream ms)
        {
            var provider = new ContentNormalizedDataHashInputProvider(project);
            provider.CreateHashInput(ms);
        }

        public static void GetFormsNormalizedDataHashInput(ExcelVbaProject project, MemoryStream ms)
        {
            var provider = new FormsNormalizedDataHashInputProvider(project);
            provider.CreateHashInput(ms);
        }

        public static void GetV3HashInput(ExcelVbaProject project, MemoryStream ms)
        {
            var provider = new V3NormalizedDataHashInputProvider(project);
            provider.CreateHashInput(ms);
        }
    }
}
