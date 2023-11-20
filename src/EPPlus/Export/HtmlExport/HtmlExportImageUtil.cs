using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal static class HtmlExportImageUtil
    {
        private static string GetClassName(string className, string optionalName)
        {
            if (string.IsNullOrEmpty(optionalName)) return optionalName;

            className = className.Trim().Replace(" ", "-");
            var newClassName = "";
            for (int i = 0; i < className.Length; i++)
            {
                var c = className[i];
                if (i == 0)
                {
                    if (c == '-' || (c >= '0' && c <= '9'))
                    {
                        newClassName = "_";
                        continue;
                    }
                }

                if ((c >= '0' && c <= '9') ||
                   (c >= 'a' && c <= 'z') ||
                   (c >= 'A' && c <= 'Z') ||
                    c >= 0x00A0)
                {
                    newClassName += c;
                }
            }
            return string.IsNullOrEmpty(newClassName) ? optionalName : newClassName;
        }

        internal static string GetPictureName(HtmlImage p)
        {
            var hash = ((IPictureContainer)p.Picture).ImageHash;
            var fi = new FileInfo(p.Picture.Part.Uri.OriginalString);
            var name = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length);

            return GetClassName(name, hash);
        }

        internal static void AddImage(HTMLElement element, HtmlExportSettings settings, HtmlImage image, object value)
        {
            if (image != null)
            {
                var imgElement = new HTMLElement(HtmlElements.Img);
                var name = GetPictureName(image);
                string imageName = GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
                imgElement.AddAttribute("alt", image.Picture.Name);
                if (settings.Pictures.AddNameAsId)
                {
                    imgElement.AddAttribute("id", imageName);
                }
                imgElement.AddAttribute("class", $"{settings.StyleClassPrefix}image-{name} {settings.StyleClassPrefix}image-prop-{imageName}");
                element.AddChildElement(imgElement);
                //writer.RenderBeginTag(HtmlElements.Img, true);
            }
        }
    }
}
