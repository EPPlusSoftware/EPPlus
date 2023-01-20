using System;
using System.Xml;

namespace OfficeOpenXml.Utils
{
    internal static class XmlReaderHelper
    {
        internal static bool ReadUntil(this XmlReader xr, int depth, params string[] tagName)
        {
            if (xr.EOF) return false;
            while ((xr.Depth == depth && Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag))) == false)
            {
                do
                {
                    xr.Read();
                    if (xr.EOF) return false;
                } while (xr.Depth != depth);
            }
            return (ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]));
        }

        /// <summary>
        /// Read file until a tag in tagName is found or EOF.
        /// This requires more careful consideration than when specifing depth.
        /// </summary>
        /// <param name="xr">Handle to xml to read data from</param>
        /// <param name="tagName">Array of tags to stop at in the order they should appear in the xml</param>
        /// <returns></returns>
        internal static bool ReadUntil(this XmlReader xr, params string[] tagName)
        {
            if (xr.EOF) return false;
            do
            {
                if (xr.EOF) return false;
                xr.Read();
            } while ((Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag))) == false);

            return (ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]));
        }
    }
}
