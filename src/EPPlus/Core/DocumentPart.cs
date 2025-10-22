/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
namespace OfficeOpenXml.Core
{
    public abstract class DocumentPart<T> where T : DocumentPart<T>
    {
        internal IDocumentPart<T> _dp;
        internal DocumentPart(IDocumentPart<T> dp)
        {
            _dp = dp;
            _dp.Load((T)this);
        }
        internal virtual void Save()
        {
            _dp.Save((T)this);
        }
        internal virtual void Remove()
        {
            _dp.Remove();
        }
    }
}