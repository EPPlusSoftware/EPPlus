using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    public class ExcelChartEx : ExcelChart
    {
        internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, eChartType? type, bool isPivot, ExcelGroupShape parent) : 
            base(drawings, node,type, isPivot, parent, "mc:AlternateContent/mc:choice/xdr:graphicFrame")
        {

        }

        internal ExcelChartEx(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml = null, ExcelGroupShape parent = null) :
            base(drawings, drawingsNode, type, topChart, PivotTableSource, chartXml, parent, "mc:AlternateContent/mc:choice/xdr:graphicFrame")
        {
        }


        internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent, "mc:AlternateContent/mc:choice/xdr:graphicFrame")
        {
        }
        //public eChartExType ChartType { get; set; }
        //public ExcelTextFont Font => throw new NotImplementedException();

        //public ExcelTextBody TextBody => throw new NotImplementedException();

        //public ExcelDrawingBorder Border => throw new NotImplementedException();

        //public ExcelDrawingEffectStyle Effect => throw new NotImplementedException();

        //public ExcelDrawingFill Fill => throw new NotImplementedException();

        //public ExcelDrawing3D ThreeD => throw new NotImplementedException();

        //public ExcelChartStyleManager StyleManager => throw new NotImplementedException();

        //public ExcelWorksheet WorkSheet => throw new NotImplementedException();

        //public XmlDocument ChartXml => throw new NotImplementedException();

        //public ExcelChartTitle Title => throw new NotImplementedException();

        //public bool HasTitle => throw new NotImplementedException();

        //public bool HasLegend => throw new NotImplementedException();

        //public ExcelChartSeries<ExcelChartSerie> Series => throw new NotImplementedException();

        //public ExcelChartAxis[] Axis => throw new NotImplementedException();

        //public ExcelChartAxis XAxis => throw new NotImplementedException();

        //public ExcelChartAxis YAxis => throw new NotImplementedException();

        //public bool UseSecondaryAxis => throw new NotImplementedException();

        //public eChartStyle Style => throw new NotImplementedException();

        //public bool RoundedCorners { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        //public bool ShowHiddenData { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        //public eDisplayBlanksAs DisplayBlanksAs { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        //public bool ShowDataLabelsOverMaximum { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        //public ExcelChartPlotArea PlotArea => throw new NotImplementedException();

        //public ExcelChartLegend Legend => throw new NotImplementedException();

        //public ExcelView3D View3D => throw new NotImplementedException();

        //public eGrouping Grouping => throw new NotImplementedException();

        //public bool VaryColors => throw new NotImplementedException();

        //public ExcelPivotTable PivotTableSource => throw new NotImplementedException();

        //ExcelPackage IPictureRelationDocument.Package => throw new NotImplementedException();

        //Dictionary<string, HashInfo> IPictureRelationDocument.Hashes => throw new NotImplementedException();

        //ZipPackagePart IPictureRelationDocument.RelatedPart => throw new NotImplementedException();

        //Uri IPictureRelationDocument.RelatedUri => throw new NotImplementedException();

        //eChartType IExcelChart.ChartType => throw new NotImplementedException();

        //public void CreatespPr()
        //{
        //    throw new NotImplementedException();
        //}

        //public void DeleteTitle()
        //{
        //    throw new NotImplementedException();
        //}

        //public void SetMandatoryProperties()
        //{
        //    throw new NotImplementedException();
        //}
    }
}
