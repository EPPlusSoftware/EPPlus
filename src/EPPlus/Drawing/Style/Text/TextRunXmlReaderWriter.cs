using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Drawing.Style.Text
{
    internal static class TextRunXmlReaderWriter
    {
        internal static RegularTextRun ReadXmlTextRun(ExcelTextFont runProperties)
        {
            runProperties.ParseAttributesFromXML();
            runProperties.ParseNodesFromXML();

            return runProperties.TextRun;
        }

        internal static ExcelDrawingTextRunCollection ReadTextRunCollection(ExcelDrawingParagraph paragraph)
        {
            var txtRunCollection = new ExcelDrawingTextRunCollection();

            var pictDoc = paragraph.DefaultRunProperties.PictureRelationDocument;

            foreach (XmlElement node in paragraph.TopNode.ChildNodes)
            {
                if (node.LocalName == "r")
                {
                    ExcelTextFont runProperties = new ExcelTextFont(pictDoc, paragraph.NameSpaceManager, node, "a:rPr", paragraph.SchemaNodeOrder);
                    var textRun = ReadXmlTextRun(runProperties);
                    txtRunCollection.Add(textRun);
                }
            }

            return txtRunCollection;
        }

        internal static void WriteTextRunCollection(ExcelDrawingTextRunCollection txtRunCollection)
        {

        }
    }
}
