using OfficeOpenXml.Drawing;
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.ImageUtils
{
    internal class ImageUrlCache
    {
        public ImageUrlCache(PictureStore pictureStore)
        {
            _pictureStore = pictureStore;
        }

        private readonly PictureStore _pictureStore;
        private readonly Dictionary<string, string> _urlCache = new Dictionary<string, string>();

        internal ImageInfo Get(string url)
        {
            if (_urlCache.ContainsKey(url))
            {
                var hash = _urlCache[url];
                if(hash==null)
                {
                    return new ImageInfo() { Uri = new Uri(url) };
                }
                return _pictureStore.GetImageInfoByHash(hash);
            }
            return null;
        }

        internal byte[] GetImageBytes(string url)
        {
            var ii = Get(url);
            if (ii == null) return null;
            return _pictureStore.GetImageBytes(ii.Uri);
        }

        internal void Add(string url, byte[] imageBytes)
        {
            if(imageBytes == null)
            {
                _urlCache[url] = null;
            }
            else
            {
                var hash = _pictureStore.GetImageHash(imageBytes);
                if (_urlCache.ContainsKey(url))
                {
                    _urlCache[url] = hash;
                }
                else
                {
                    _urlCache.Add(url, hash);
                }
            }
        }
    }
}
