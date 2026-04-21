using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
//{
//    internal class StyleDrawing : IStyleExport
//    {
//        //ExcelChartStyleEntry _style;

//        //public bool HasStyle
//        //{
//        //    get { return _style.HasFill; }
//        //}

//        //public string StyleKey { get; } = null; /*{ get { return _style.Id; } }*/

//        //public IFill Fill { get; } = null;
//        //public IFont Font { get; } = null;
//        //public IBorder Border { get; } = null;
//        //public INumberFormat NumberFormat { get; } = null;

//        ////Charts never have a checkbox. Break out of baseClass?
//        //bool IStyleExport.CheckBox => false;

//        //public int StyleId;

//        //public StyleDrawing(ExcelChartStyleEntry style, int styleId)
//        //{
//        //    _style = style;
//        //    StyleId = styleId;

//        //    if (style.HasFill)
//        //    {
//        //        Fill = new FillDrawing(style.Fill);
//        //    }
//        //    if (style.HasTextBody)
//        //    {
//        //        //not implemented yet
//        //        if (style.HasRichText)
//        //        {
//        //            if(style.HasTextRun)
//        //            {
//        //                //style.FontReference
//        //            }
//        //        }
//        //        //Font = new FontDxf(style.Font);
//        //    }
//        //    if(style.FontReference != null)
//        //    {
//        //        //not implemented yet
//        //    }
//        //    if (style.HasBorder)
//        //    {
//        //        //style.Border.
//        //        //var lineStyle = style.Border.LineStyle;
//        //        //Border.Top.Style = Style.ExcelBorderStyle.
//        //        //Border = new FillDrawingBasic(style.Border.Fill);
//        //    }
//        //    //if (style. != null && style.NumberFormat.HasValue)
//        //    //{
//        //    //    NumberFormat = new NumberFormatDxf(style.NumberFormat, styleId);
//        //    //}

//        //    //CheckBox = false;
//        //}
//    }
//}
