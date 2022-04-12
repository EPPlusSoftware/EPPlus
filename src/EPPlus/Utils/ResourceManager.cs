/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Packaging.Ionic.Zip;
using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal class StyleResourceManager
    {
        internal static string GetItem(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream=assembly.GetManifestResourceStream("OfficeOpenXml.resources.DefaultChartStyles.ecs");

            using (stream)
            {
                var zipStream = new ZipInputStream(stream);
                ZipEntry entry;
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (entry.IsDirectory || !entry.FileName.EndsWith(".xml") || entry.UncompressedSize <= 0) continue;

                    var fileName = new FileInfo(entry.FileName).Name;

                    if(fileName.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    {
                        //Extract and set
                        var content = new byte[entry.UncompressedSize];
                        var size = zipStream.Read(content, 0, (int)entry.UncompressedSize);
                        return Encoding.UTF8.GetString(content);
                    }
                }
            }
            return null;
        }
    }
}
