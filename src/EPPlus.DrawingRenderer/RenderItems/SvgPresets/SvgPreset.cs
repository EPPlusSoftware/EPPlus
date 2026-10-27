using EPPlus.DrawingRenderer.Utils;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.DrawingRenderer.RenderItems.SvgPresets
{
    abstract internal class SvgPreset<T> where T : RenderItem
    {
        internal double width;

        internal double height;


        //internal SvgPreset() 
        //{
        //    //width = item.Width;
        //    //height = item.Height;
        //}

        //public virtual void SetFillColor(Color color)
        //{
        //    this.Style.FillColor = "#" + color.ToColorString();
        //}
    }
}
