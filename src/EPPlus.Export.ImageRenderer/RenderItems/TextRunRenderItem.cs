using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems
{
    internal abstract class TextRunRenderItem : RenderItem
    {
        public override SvgItemType Type => SvgItemType.Text;

        internal TextRunRenderItem(ExcelParagraphTextRunBase textRun) 
        {
            
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            throw new NotImplementedException();
        }
    }
}
