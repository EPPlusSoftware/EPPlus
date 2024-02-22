using System;
using System.Collections.Generic;
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
                } while (!(xr.Depth == depth && xr.NodeType == XmlNodeType.Element));
            }
            return xr.NodeType == XmlNodeType.Element && ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]);
        }

        /// <summary>
        /// Read file until a tag in tagName is found or EOF.
        /// This requires more careful consideration than when specifing depth.
        /// As it will exit on endnodes and continue directly to end of file if nothing is found.
        /// </summary>
        /// <param name="xr">Handle to xml to read data from</param>
        /// <param name="tagName">Array of tags to stop at in the order they should appear in the xml</param>
        /// <returns>false if EOF or found end tag. True if found tag of element type</returns>
        internal static bool ReadUntil(this XmlReader xr, params string[] tagName)
        {
            if (xr.EOF) return false;
            do
            {
                if (xr.EOF) return false;
                xr.Read();
            } while ((Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag))) == false);

            return xr.NodeType == XmlNodeType.Element && ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]);
        }

        /// <summary>
        /// Returns false when end or EOF is found
        /// </summary>
        /// <param name="xr"></param>
        /// <param name="child"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        internal static bool ReadUntilNextChildOrParentEnd(this XmlReader xr, string child, string parent)
        {
            if (xr.EOF) return false;
            do
            {
                if (xr.EOF) return false;
                xr.Read();
            } while (StartOrEndFound(xr.LocalName, child, parent, xr.NodeType) == false);

            return xr.NodeType == XmlNodeType.Element;
        }

        private static bool StartOrEndFound(string localName, string start, string end, XmlNodeType nodeType)
        {
            if(ConvertUtil._invariantCompareInfo.IsSuffix(localName, start))
            {
                return nodeType == XmlNodeType.Element;
            }
            else if(ConvertUtil._invariantCompareInfo.IsSuffix(localName, end))
            {
                return nodeType == XmlNodeType.EndElement;
            }
            return false;
        }

        //internal static bool Read(this XmlReader xr, params string[] tagName)
        //{
        //    if (xr.EOF) return false;
        //    do
        //    {
        //        if (xr.EOF) return false;
        //        xr.Read();
        //    } while ((Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag))) == false);

        //    return xr.NodeType == XmlNodeType.Element && ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]);
        //}


        /// <summary>
        /// Read file until a start tag in tagName is found or EOF.
        /// This requires more careful consideration than when specifing depth.
        /// As it will only exit on element nodes and continue directly to end of file if nothing is found.
        /// </summary>
        /// <param name="xr">Handle to xml to read data from</param>
        /// <param name="tagName">Array of tags to stop at in the order they should appear in the xml</param>
        /// <returns></returns>
        internal static bool ReadUntilElement(this XmlReader xr, params string[] tagName)
        {
            if (xr.EOF) return false;
            do
            {
                if (xr.EOF) return false;
                xr.Read();
            } while ((Array.Exists(tagName, tag => ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag) && xr.NodeType == XmlNodeType.Element)) == false);

            return xr.NodeType == XmlNodeType.Element && ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]);
        }

        internal static bool ReadUntil(this XmlReader xr, int depth, Dictionary<string, int> nodeOrder, string tag)
        {
            if (xr.EOF == false && nodeOrder.TryGetValue(tag, out int tagIx))
            {
                if (nodeOrder.TryGetValue(xr.LocalName, out int currentNodeIx))
                {
                    while ((xr.Depth == depth && currentNodeIx < tagIx))
                    {
                        do
                        {
                            xr.Read();

                            if (xr.EOF)
                            {
                                return false;
                            }
                        } while (xr.Depth != depth);
                        if (!nodeOrder.TryGetValue(xr.LocalName, out currentNodeIx))
                        {
                            return false;
                        }
                    }
                    return xr.NodeType == XmlNodeType.Element && ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag);
                }
            }
            return false;
        }

    }
}
