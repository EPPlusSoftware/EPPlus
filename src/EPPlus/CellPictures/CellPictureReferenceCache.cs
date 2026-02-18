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
using OfficeOpenXml.RichData;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.CellPictures
{
    internal class CellPictureReferenceCache
    {
        private readonly Dictionary<string, CellPictureReference> _referenceCache = new Dictionary<string, CellPictureReference>();
        public EventHandler<LastReferenceRemovedEventArgs> LastReferenceRemoved;

        private void OnLastReferenceRemoved(uint vmId)
        {
            LastReferenceRemoved?.Invoke(this, new LastReferenceRemovedEventArgs(vmId));
        }

        public static PictureCacheKey CreateKey(ExcelCellPicture picture)
        {
            if (picture.PictureType == ExcelCellPictureTypes.LocalImage)
            {
                return new LocalImageCacheKey(picture);
            }
            else if (picture.PictureType == ExcelCellPictureTypes.WebImage)
            {
                return new WebPictureCacheKey(picture);
            }
            else
            {
                throw new ArgumentException("Invalid pictureType: " + picture.PictureType.ToString());
            }
        }

        public bool Contains(PictureCacheKey pictureCacheKey, out uint vmId)
        {
            vmId = uint.MaxValue;
            var key = pictureCacheKey.Key;
            var result = _referenceCache.ContainsKey(key);
            if(result)
            {
                vmId = _referenceCache[key].VmId; 
            }

            return result;
        }

        public bool Contains(PictureCacheKey pictureCacheKey)
        {
            var key = pictureCacheKey.Key;
            return _referenceCache.ContainsKey(key);
        }

        public int GetNumberOfReferences(PictureCacheKey pictureCacheKey)
        {
            int refNum = 0;
            var key = pictureCacheKey.Key;
            if (_referenceCache.ContainsKey(key))
            {
                refNum = _referenceCache[key].NumberOfReferences;
            }
            return refNum;
        }

        public void Add(PictureCacheKey pictureCacheKey, uint vmId)
        {
            var key = pictureCacheKey.Key;
            if (!_referenceCache.ContainsKey(key) )
            {
                _referenceCache[key] = new CellPictureReference(vmId);
            }
            _referenceCache[key].AddReference();
        }

        public bool Remove(PictureCacheKey pictureCacheKey)
        {
            var key = pictureCacheKey.Key;
            if ( _referenceCache.ContainsKey(key) )
            {
                var item = _referenceCache[key];
                item.RemoveReference();
                if(item.NumberOfReferences == 0)
                {
                    OnLastReferenceRemoved(item.VmId);
                    _referenceCache.Remove(key);
                }
                return true;
            }
            return false;
        }

        public void AddReference(PictureCacheKey pictureCacheKey)
        {
            var key = pictureCacheKey.Key;
            if(_referenceCache.ContainsKey(key))
            {
                _referenceCache[key].AddReference();
            }
        }

        public void RemoveReference(PictureCacheKey pictureCacheKey, out int numberOfReferencesLeft)
        {
            numberOfReferencesLeft = 0;
            var key = pictureCacheKey.Key;
            if (_referenceCache.ContainsKey(key))
            {
                _referenceCache[key].RemoveReference();
                numberOfReferencesLeft = _referenceCache[key].NumberOfReferences;
            }
        }


    }
}
