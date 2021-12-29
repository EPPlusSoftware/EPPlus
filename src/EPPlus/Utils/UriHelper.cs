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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal class UriHelper
    {
        internal static Uri ResolvePartUri(Uri sourceUri, Uri targetUri)
        {
            if (targetUri.OriginalString.StartsWith("/", StringComparison.OrdinalIgnoreCase) || targetUri.OriginalString.Contains("://"))
            {
                return targetUri;
            }
            string[] source = sourceUri.OriginalString.Split('/');
            string[] target = targetUri.OriginalString.Split('/');

            int t = target.Length - 1;
            int s;
            if(sourceUri.OriginalString.EndsWith("/")) //is the source a directory?
            {
                s = source.Length-1;
            }
            else
            {
                s=source.Length-2;
            }

            string file = target[t--];

            while (t >= 0)
            {
                if (target[t] == ".")
                {
                    break;
                }
                else if (target[t] == "..")
                {
                    s--;
                    t--;
                }
                else
                {
                    file = target[t--] + "/" + file;
                }
            }
            if (s >= 0)
            {
                for(int i=s;i>=0;i--)
                {
                    file = source[i] + "/" + file;
                }
            }
            return new Uri(file,UriKind.RelativeOrAbsolute);
        }

        internal static Uri GetRelativeUri(Uri WorksheetUri, Uri uri)
        {
            string[] source = WorksheetUri.OriginalString.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string[] target = uri.OriginalString.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

            int slen;
            if (WorksheetUri.OriginalString.EndsWith("/"))
            {
                slen = source.Length;
            }
            else
            {
                slen = source.Length-1;
            }
            int i = 0;
            while (i < slen && i < target.Length && source[i] == target[i])
            {
                i++;
            }

            string dirUp="";
            for (int s = i; s < slen; s++)
            {
                dirUp += "../";
            }
            string file = "";
            for (int t = i; t < target.Length; t++)
            {                
                file += (file==""?"":"/") + target[t];
            }
            return new Uri(dirUp+file,UriKind.Relative);
        }
    }
}
