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
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Vba.ContentHash
{
    internal class ContentNormalizedDataHashInputProvider : ContentHashInputProvider
    {
        public ContentNormalizedDataHashInputProvider(ExcelVbaProject project) : base(project)
        {
        }

        protected override void CreateHashInputInternal(MemoryStream ms)
        {
            GetContentHash(ms);
        }

        private void GetContentHash(MemoryStream ms)
        {
            //MS-OVBA 2.4.2.1
            BinaryWriter bw = new BinaryWriter(ms);
            bw.Write(HashEncoding.GetBytes(Project.Name));
            bw.Write(HashEncoding.GetBytes(Project.Constants));
            foreach (var reference in Project.References)
            {
                if (reference.ReferenceRecordID == 0x0D)
                {
                    bw.Write((byte)0x7B);
                }
                else if (reference.ReferenceRecordID == 0x0E)
                {
                    foreach (byte b in BitConverter.GetBytes((uint)reference.Libid.Length))  //Length will never be an UInt with 4 bytes that aren't 0 (> 0x00FFFFFF), so no need for the rest of the properties.
                    {
                        if (b != 0)
                        {
                            bw.Write(b);
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }
            foreach (var module in Project.Modules)
            {
                var lines = module.Code.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (!line.StartsWith("attribute", StringComparison.OrdinalIgnoreCase))
                    {
                        bw.Write(HashEncoding.GetBytes(line));
                    }
                }
            }
        }
    }
}
