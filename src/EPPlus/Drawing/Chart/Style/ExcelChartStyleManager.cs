/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// Manages styles for a chart
    /// </summary>
    public class ExcelChartStyleManager : XmlHelper
    {
        internal readonly ExcelChart _chart;
        private readonly ExcelThemeManager _theme;
        private static bool _hasLoadedLibraryFiles=false;
        internal ExcelChartStyleManager(XmlNamespaceManager nameSpaceManager, ExcelChart chart) : base(nameSpaceManager)
        {
            _chart = chart;
            LoadStyleAndColors(chart);
            if (StylePart != null)
            {
                Style = new ExcelChartStyle(nameSpaceManager, StyleXml.DocumentElement, this);
            }
            if (ColorsPart != null)
            {
                ColorsManager = new ExcelChartColorsManager(nameSpaceManager, ColorsXml.DocumentElement);
            }
            _theme = chart.WorkSheet.Workbook.ThemeManager;
        }
        /// <summary>
        /// A library where chart styles can be loaded for easier access.
        /// EPPlus loads most buildin styles into this collection.
        /// </summary>
        public static Dictionary<int, ExcelChartStyleLibraryItem> StyleLibrary = new Dictionary<int, ExcelChartStyleLibraryItem>();
        /// <summary>
        /// A library where chart color styles can be loaded for easier access
        /// </summary>
        public static Dictionary<int, ExcelChartStyleLibraryItem> ColorsLibrary = new Dictionary<int, ExcelChartStyleLibraryItem>();
        /// <summary>
        /// Creates an empty style and color for chart, ready to be customized 
        /// </summary>
        public void CreateEmptyStyle(eChartStyle fallBackStyle=eChartStyle.Style2)
        {
            if (fallBackStyle == eChartStyle.None)
            {
                throw new InvalidOperationException("The chart must have a style. Please set the charts Style property to a value different than None or Call LoadStyleXml with the fallBackStyle parameter");
            }

            var p = _chart.WorkSheet.Workbook._package.ZipPackage;
            var id = CreateStylePart(p);
            StyleXml = new XmlDocument();
            StyleXml.LoadXml(GetStartStyleXml(id));
            StyleXml.Save(StylePart.GetStream());
            Style = new ExcelChartStyle(NameSpaceManager, StyleXml.DocumentElement, this);
            _chart.InitChartTheme((int)fallBackStyle);

            CreateColorXml(p);
        }

        private void CreateColorXml(ZipPackage p)
        {
            CreateColorPart(p);
            ColorsXml = new XmlDocument();
            ColorsXml.LoadXml(GetStartColorXml());
            ColorsXml.Save(ColorsPart.GetStream());

            ColorsManager = new ExcelChartColorsManager(NameSpaceManager, ColorsXml.DocumentElement);
        }
        #region LoadStyles
        /// <summary>
        /// Loads the default chart style library from the internal resource library.
        /// Loads styles, colors and the default theme.
        /// </summary>
        public static void LoadStyles()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var defaultStyleLibrary = assembly.GetManifestResourceStream("OfficeOpenXml.resources.DefaultChartStyles.ecs");

            LoadStyles(defaultStyleLibrary);
        }
        /// <summary>
        /// Load all chart style library files (*.ecs) into memory from the supplied directory
        /// </summary>
        /// <param name="directory">Load all *.ecs files from the directory</param>
        /// <param name="clearLibrary">If true, clear the library before load.</param>
        public static void LoadStyles(DirectoryInfo directory, bool clearLibrary=true)
        {
            if (clearLibrary)
            {
                StyleLibrary.Clear();
            }
            foreach (var ecsFile in directory.GetFiles("*.ecs"))
            {
                LoadStyles(ecsFile, false);
            }
        }
        /// <summary>
        /// Load a single chart style library file (*.ecs) into memory
        /// </summary>
        /// <param name="ecsFile">The file to load</param>
        /// <param name="clearLibrary">If true, clear the library before load.</param>
        public static void LoadStyles(FileInfo ecsFile, bool clearLibrary = true)
        {
            using (var fs = ecsFile.Open(FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                LoadStyles(fs, clearLibrary, fs.Name);
            }
        }
        /// <summary>
        /// Load a single chart style library stream into memory from the supplied directory
        /// </summary>
        /// <param name="stream">The stream to load</param>
        /// <param name="clearLibrary">If true, clear the library before load.</param>
        public static void LoadStyles(Stream stream, bool clearLibrary = true)
        {
            LoadStyles(stream, clearLibrary, "The stream");
        }
        private static void LoadStyles(Stream stream, bool clearLibrary, string filename)
        {
            if (clearLibrary)
            {
                StyleLibrary.Clear();
            }
            try
            {
                using (stream)
                {
                    var zipStream = new ZipInputStream(stream);
                    ZipEntry entry;
                    while ((entry = zipStream.GetNextEntry()) != null)
                    {
                        if (entry.IsDirectory || !entry.FileName.EndsWith(".xml") || entry.UncompressedSize <= 0) continue;

                        var name = new FileInfo(entry.FileName).Name;
                        int id;
                        try
                        {
                            if (name.StartsWith("colors", StringComparison.InvariantCultureIgnoreCase))
                            {
                                id = int.Parse(name.Substring(6, name.Length - 10));
                                if (ColorsLibrary.ContainsKey(id)) continue;
                            }
                            else if (name.StartsWith("style", StringComparison.InvariantCultureIgnoreCase))
                            {
                                id = int.Parse(name.Substring(5, name.Length - 9));
                                if (StyleLibrary.ContainsKey(id)) continue;
                            }
                            else if (name.Equals("defaulttheme.xml", StringComparison.InvariantCultureIgnoreCase))
                            {
                                var themeXml = UncompressEntry(zipStream, entry);
                                ExcelThemeManager._defaultTheme = themeXml;
                                continue;
                            }
                            else
                            {
                                throw (new InvalidDataException($"{filename} contains a the file {entry.FileName}, with an invalid filename. Please make sure files in the library are named Colors[id].xml or style[id].xml, where [id] is replaced by the id to access the style in the library"));
                            }
                        }
                        catch
                        {
                            throw (new InvalidDataException($"{filename} contains a the file {entry.FileName}, with an invalid filename. Please make sure files in the library are named Colors[id].xml or style[id].xml, where [id] is replaced by the id to access the style in the library"));
                        }


                        //Extract and set
                        var uncompressedContent = UncompressEntry(zipStream, entry);
                        var item = new ExcelChartStyleLibraryItem() { Id = id, XmlString = uncompressedContent };
                        if (name[0] == 'c') //Colors
                        {
                            ColorsLibrary.Add(item.Id, item);
                        }
                        else
                        {
                            StyleLibrary.Add(item.Id, item);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                if(ex is InvalidDataException)
                {
                    throw;
                }
                else
                {
                    throw (new InvalidDataException($"{filename} has an invalid format.", ex));
                }
            }
        }

        private static string UncompressEntry(ZipInputStream zipStream, ZipEntry entry)
        {
            var content = new byte[entry.UncompressedSize];
            var size = zipStream.Read(content, 0, (int)entry.UncompressedSize);
            return Encoding.UTF8.GetString(content);            
        }
        #endregion
        /// <summary>
        /// Loads a chart style xml file, and applies the style.
        /// </summary>
        /// <param name="styleXml">The chart style xml document</param>
        /// <param name="colorXml">The chart color xml document</param>
        /// <returns>The new Id of the Style loaded</returns>
        /// <remarks>
        /// This is the style.xml and colors.xml related to the chart.xml inside a package or chart template, 
        /// e.g \xl\charts\chart1.xml
        ///     \xl\charts\style1.xml 
        ///     \xl\charts\colors1.xml
        /// </remarks>
        public int LoadStyleXml(XmlDocument styleXml, XmlDocument colorXml = null)
        {
            if (_chart.Style == eChartStyle.None)
            {
                _chart.Style = eChartStyle.Style2;
            }
            return LoadStyleXml(StyleXml, _chart.Style, colorXml);
        }
        /// <summary>
        /// Loads a crtx file and applies it to the chart. Crtx files are exported from a Spreadsheet Application like Excel.
        /// Loading a template will only apply the styles to the chart, not change settings for the chart.
        /// Please use the <c>AddChartFromTemplate</c> method to add a chart from a template file.
        /// </summary>
        /// <param name="crtxFile">A crtx file</param>
        /// <seealso cref="ExcelDrawings.AddChartFromTemplate(FileInfo, string)"/>
        public void LoadTemplateStyles(FileInfo crtxFile)
        {
            if (!crtxFile.Exists)
            {
                throw (new FileNotFoundException($"{crtxFile.FullName} cannot be found."));
            }
            FileStream fs = null;
            try
            {
                fs = crtxFile.Open(FileMode.Open, FileAccess.Read, FileShare.Read);
                LoadTemplateStyles(fs, crtxFile.Name);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }
        }
        /// <summary>
        /// Loads a crtx file and applies it to the chart. Crtx files are exported from a Spreadsheet Application like Excel.
        /// Loading a template will only apply the styles to the chart, not change settings for the chart, override themes etc.
        /// Please use the <c>AddChartFromTemplate</c> method to add a chart from a template file.
        /// </summary>
        /// <param name="crtxStream">A stream containing a crtx file</param>
        /// <seealso cref="ExcelDrawings.AddChartFromTemplate(Stream, string)"/>
        public void LoadTemplateStyles(Stream crtxStream)
        {
            LoadTemplateStyles(crtxStream, "The crtx stream");
        }

        private void LoadTemplateStyles(Stream crtxStream, string name)
        {
            CrtxTemplateHelper.LoadCrtx(crtxStream, out XmlDocument chartXml, out XmlDocument styleXml, out XmlDocument colorsXml, out _, name);
            if (!(styleXml == null && colorsXml == null))
            {
                //TODO:Add theme override rel to chart.
                //TODO:Add all settings for chart.xml.
                LoadStyleXml(styleXml, eChartStyle.Style2, colorsXml);
                ApplyStyles();
            }
            else
            {
                throw new InvalidDataException("Crtx file is corrupt.");
            }
        }

        /// <summary>
        /// Loads a chart style xml file, and applies the style.
        /// </summary>
        /// <param name="fallBackStyle">The build in style to fall back on</param>
        /// <param name="styleXml">The chart style xml document</param>
        /// <param name="colorsXml">The chart colord xml document</param>
        /// <returns>The id of the Style loaded</returns>
        public int LoadStyleXml(XmlDocument styleXml, eChartStyle fallBackStyle, XmlDocument colorsXml=null)
        {
            LoadStyleAndColorsXml(styleXml, fallBackStyle, colorsXml);

            if (_chart._isChartEx)
            {
                ApplyStylesEx();
            }
            else
            {
                ApplyStyles();
            }

            return Style.Id;
        }   
        internal void LoadStyleAndColorsXml(XmlDocument styleXml, eChartStyle fallBackStyle, XmlDocument colorsXml)
        {
            if (fallBackStyle == eChartStyle.None)
            {
                throw new ArgumentException("fallBackStyle", "fallBackStyle can't be None");
            }
            if (_chart.Style != eChartStyle.None && _chart.Style != fallBackStyle)
            {
                _chart.Style = fallBackStyle;
            }

            if (styleXml == null || styleXml.DocumentElement == null || styleXml.DocumentElement.LocalName != "chartStyle" || styleXml.DocumentElement.ChildNodes.Count != 31)
            {
                throw new ArgumentException("xml", "StyleXml is null or not in the correct format");
            }

            if (StylePart == null)
            {
                CreateStylePart(_chart.WorkSheet.Workbook._package.ZipPackage);
            }

            StyleXml = styleXml;
            StyleXml.Save(StylePart.GetStream(FileMode.CreateNew));
            Style = new ExcelChartStyle(NameSpaceManager, StyleXml.DocumentElement, this);

            if (colorsXml == null)
            {
                colorsXml = new XmlDocument();
                colorsXml.LoadXml(GetStartColorXml());
            }

            LoadColorXml(colorsXml);
            if (_chart._isChartEx==false)
            {
                _chart.InitChartTheme((int)fallBackStyle);
            }
        }
        /// <summary>
        /// Loads a theme override xml document for the chart.
        /// </summary>
        /// <param name="themePart">The theme part</param>
        internal void LoadThemeOverrideXml(ZipPackagePart themePart)
        {
            var rel = CreateThemeOverridePart(_chart.WorkSheet.Workbook._package.ZipPackage, themePart);
            ThemeOverride = new ExcelThemeOverride(_chart, rel);
        }
        /// <summary>
        /// Applies a preset chart style loaded into the StyleLibrary to the chart.
        /// </summary>
        /// <param name="style">The style to use</param>
        /// <seealso cref="SetChartStyle(int, int?)"/>
        public void SetChartStyle(ePresetChartStyle style)
        {
            SetChartStyle(style, ePresetChartColors.ColorfulPalette1);
        }
        /// <summary>
        /// Applies a preset chart style loaded into the StyleLibrary to the chart.
        /// </summary>
        /// <param name="style">The style to use</param>
        /// <seealso cref="SetChartStyle(int, int?)"/>
        public void SetChartStyle(ePresetChartStyleMultiSeries style)
        {
            SetChartStyle(style, ePresetChartColors.ColorfulPalette1);
        }
        /// <summary>
        /// Applies a preset chart style loaded into the StyleLibrary to the chart. 
        /// This enums matches Excel's styles for single series for common scenarios. 
        /// Excel changes chart styles depending on many parameters, like number of series, axis type and more, so it will not always match the number in Excel.       
        /// To be certain of getting the correct style use the chart style number of the style you want to apply
        /// </summary>
        /// <param name="style">The preset style to use</param>
        /// <param name="colors">The preset color scheme to use</param>
        /// <seealso cref="SetChartStyle(int, int?)"/>
        public void SetChartStyle(ePresetChartStyle style, ePresetChartColors colors)
        {
            SetChartStyle((int)style, (int)colors);
        }
        /// <summary>
        /// Applies a preset chart style loaded into the StyleLibrary to the chart.
        /// This enums matches Excel's styles for multiple series for common scenarios. 
        /// Excel changes chart styles depending on many parameters, like number of series, axis type and more, so it will not always match the number in Excel.       
        /// To be certain of getting the correct style use the chart style number of the style you want to apply.
        /// </summary>
        /// <param name="style">The preset style to use</param>
        /// <param name="colors">The preset color scheme to use</param>
        /// <seealso cref="SetChartStyle(int, int?)"/>
        public void SetChartStyle(ePresetChartStyleMultiSeries style, ePresetChartColors colors)
        {
            SetChartStyle((int)style, (int)colors);
        }
        /// <summary>
        /// Applies a chart style loaded into the StyleLibrary to the chart.
        /// </summary>
        /// <param name="style">The chart style id to use</param>
        /// <param name="colors">The preset color scheme id to use. Null means </param>
        /// <seealso cref="SetChartStyle(ePresetChartStyle)"/>
        public void SetChartStyle(int style, int? colors = (int)ePresetChartColors.ColorfulPalette1)
        {
            if (_hasLoadedLibraryFiles == false && StyleLibrary.Count == 0)
            {
                LoadStyles();
            }

            if (!StyleLibrary.ContainsKey(style))
            {
                if(Enum.IsDefined(typeof(ePresetChartColors), style))
                {
                    throw (new KeyNotFoundException($"Style {(ePresetChartStyle)style} ({style}) cannot be found in the StyleLibrary. Please load it into the StyleLibrary."));
                }
                else
                {
                    throw (new KeyNotFoundException($"Style {style} cannot be found in the StyleLibrary. Please load it into the StyleLibrary."));
                }
            }

            if (colors.HasValue && !ColorsLibrary.ContainsKey(colors.Value))
            {
                if (Enum.IsDefined(typeof(ePresetChartColors), colors.Value))
                {
                    throw (new KeyNotFoundException($"Colors scheme {(ePresetChartColors)colors} ({colors}) cannot be found in the ColorsLibrary. Please load it into the ColorsLibrary."));
                }
                else
                {
                    throw (new KeyNotFoundException($"Colors scheme {colors} cannot be found in the ColorsLibrary. Please load it into the ColorsLibrary."));
                }
            }
            _chart.Style = eChartStyle.None;
            if (colors.HasValue)
            {
                LoadStyleXml(StyleLibrary[style].XmlDocument, eChartStyle.Style2, ColorsLibrary[colors.Value].XmlDocument);
            }
            else
            {
                LoadStyleXml(StyleLibrary[style].XmlDocument, eChartStyle.Style2);
            }
        }
        /// <summary>
        /// Load a color xml documents
        /// </summary>
        /// <param name="colorXml">The color xml</param>
        public void LoadColorXml(XmlDocument colorXml)
        {
            if (colorXml == null || colorXml.DocumentElement == null || colorXml.DocumentElement.LocalName != "colorStyle" || colorXml.DocumentElement.ChildNodes.Count == 0)
            {
                throw new ArgumentException("xml", "ColorXml is null or not in the correct format");
            }

            if (ColorsPart == null)
            {
                CreateColorPart(_chart.WorkSheet.Workbook._package.ZipPackage);
            }

            ColorsXml = colorXml;
            var stream = ColorsPart.GetStream(FileMode.CreateNew);
            ColorsXml.Save(stream);
            
            ColorsManager = new ExcelChartColorsManager(NameSpaceManager, ColorsXml.DocumentElement);
        }
        /// <summary>
        /// Apply the chart and color style to the chart.
        /// <seealso cref="Style"/>
        /// <seealso cref="ColorsManager"/>
        /// </summary>
        public void ApplyStyles()
        {
            //Make sure we have a theme
            if (_theme.CurrentTheme == null)
            {
                _theme.CreateDefaultTheme();
            }

            if (_chart._topChart!=null)
            {
                throw (new InvalidOperationException("Please style the parent chart only"));
            }
            if (_chart.VaryColors)
            {
                GenerateDataPoints();
            }
            ApplyStyle(_chart, Style.ChartArea);

            //Plotarea
            if (_chart.IsType3D())
            {
                ApplyStyle(_chart.PlotArea, Style.PlotArea3D);
                ApplyStyle(_chart.Floor, Style.Floor);
                ApplyStyle(_chart.SideWall, Style.Wall);
                ApplyStyle(_chart.BackWall, Style.Wall);                
            }
            else
            {
                ApplyStyle(_chart.PlotArea, Style.PlotArea);
            }

            //Title
            if (_chart.HasTitle)
            {
                ApplyStyle(_chart.Title, Style.Title);
            }

            if (_chart.PlotArea.DataTable!=null)
            {
                ApplyStyle(_chart.PlotArea.DataTable, Style.DataTable);
            }

            ApplyDataLabels();
            
            if (_chart.HasLegend)
            {
                ApplyStyle(_chart.Legend, Style.Legend);
                if(!_chart._isChartEx)
                {
                    if (_chart.Legend._entries != null)
                    {
                        foreach (ExcelChartLegendEntry e in _chart.Legend._entries)
                        {
                            if (e.HasValue)
                            {
                                ApplyStyleFont(Style.Legend, e.Index, e, 0);
                            }
                        }
                    }
                }
            }
            
            if(_chart is ExcelStandardChartWithLines lineChart)
            {
                if (!(lineChart.DropLine is null)) ApplyStyle(lineChart.DropLine, Style.DropLine);
                if (!(lineChart.HighLowLine is null)) ApplyStyle(lineChart.HighLowLine, Style.HighLowLine);
                if (!(lineChart.UpBar is null)) ApplyStyle(lineChart.UpBar, Style.UpBar);
                if (!(lineChart.DownBar is null)) ApplyStyle(lineChart.DownBar, Style.DownBar);
            }

            ApplyAxis();
            ApplySeries();
        }
        /// <summary>
        /// Apply the chart and color style to the chart.
        /// <seealso cref="Style"/>
        /// <seealso cref="ColorsManager"/>
        /// </summary>
        public void ApplyStylesEx()
        {
            //Make sure we have a theme
            if (_theme.CurrentTheme == null)
            {
                _theme.CreateDefaultTheme();
            }

            //Title
            if (_chart.HasTitle && _chart.Title.TopNode.HasChildNodes)
            {
                ApplyStyle(_chart.Title, Style.Title);
            }

            if (_chart.HasLegend && _chart.Legend.TopNode.HasChildNodes)
            {
                ApplyStyle(_chart.Legend, Style.Legend);
            }

            ApplyAxis();
        }
        private void GenerateDataPoints()
        {            
            foreach (var serie in _chart.Series)
            {
                GenerateDataPointsSerie(serie);
            }
        }

        private void GenerateDataPointsSerie(ExcelChartSerie serie)
        {
            if(serie is IDrawingChartDataPoints dtpSerie)
            {
                int points;
                if(_chart.PivotTableSource==null)
                {
                    var address = _chart.WorkSheet.Workbook.GetRange(_chart.WorkSheet, serie.Series);
                    if (address == null) return;
                    points = address.Rows == 1 ? address.Columns : address.Rows;
                    
                }
                else
                {
                    points = 48;
                }
                for (var i = 0; i < points; i++)
                {
                    if (!dtpSerie.DataPoints.ContainsKey(i))
                    {
                        dtpSerie.DataPoints.AddDp(i, "0000000D-5D51-4ADD-AFBE-74A932E24C89");
                    }
                }
            }
        }

        private void ApplyDataLabels()
        {
            if (_chart is IDrawingDataLabel dataLabel)
            {
                if (dataLabel.HasDataLabel)
                {
                    ApplyStyle(dataLabel.DataLabel, Style.DataLabel);
                }
                foreach (IDrawingSerieDataLabel serie in _chart.Series)
                {
                    if (serie.HasDataLabel)
                    {
                        ApplyDataLabelSerie(serie.DataLabel);
                    }
                }
            }
        }

        private void ApplyDataLabelSerie(ExcelChartSerieDataLabel dataLabel)
        {
            ApplyStyle(dataLabel, Style.DataLabel);
            foreach (var lbl in dataLabel.DataLabels)
            {
                ApplyStyle(lbl, Style.DataLabel);
            }
        }

        private void ApplyAxis()
        {
            foreach (var axis in _chart.Axis)
            {
                ExcelChartStyleEntry currStyle;
                if (axis.AxisType == eAxisType.Cat ||
                    axis.AxisType == eAxisType.Date)
                {
                    currStyle = Style.CategoryAxis;

                }
                else if (axis.AxisType == eAxisType.Serie)
                {
                    currStyle = Style.SeriesAxis;
                }
                else
                {
                    currStyle = Style.ValueAxis;
                }
                
                if (_chart._isChartEx == false || axis._title != null)
                {
                    ApplyStyle(axis, currStyle);
                }

                if (axis.HasMajorGridlines)
                {
                    ApplyStyleBorder(axis.MajorGridlines, Style.GridlineMajor, 0, 0);
                    ApplyStyleEffect(axis.MajorGridlineEffects, Style.GridlineMajor, 0, 0);
                }
                if (axis.HasMinorGridlines)
                {
                    ApplyStyleBorder(axis.MinorGridlines, Style.GridlineMinor, 0, 0);
                    ApplyStyleEffect(axis.MinorGridlineEffects, Style.GridlineMajor, 0, 0);
                }
            }
        }

        internal void ApplySeries()
        {   
            foreach (var chart in _chart.PlotArea.ChartTypes)
            {
                var dataPoint = GetDataPointStyle(chart);
                var applyFill = (!chart.IsTypeLine() || chart.ChartType == eChartType.Line3D || chart.ChartType == eChartType.XYScatter);   //Lines have no fill, except Line3D
                int serieNo = 0;
                foreach (ExcelChartSerie serie in chart.Series)
                {
                    //Note: Datalabels are applied in the ApplyDataLabels method
                    //Marker
                    var applyBorder = !(chart.IsTypeStock() && serie.Border.Width==0);
                    ApplyStyle(serie, dataPoint, serieNo, chart.Series.Count, applyFill, applyBorder);
                    if (serie is IDrawingChartMarker serieMarker && serieMarker.HasMarker())     //Applies to Line and Scatterchart series
                    {
                        ApplyStyle(serieMarker.Marker, Style.DataPointMarker, serieNo, chart.Series.Count);
                        serieMarker.Marker.Size = Style.DataPointMarkerLayout.Size;
                        if (Style.DataPointMarkerLayout.Style != eMarkerStyle.None)
                        {
                            serieMarker.Marker.Style = Style.DataPointMarkerLayout.Style;
                        }
                    }
                    
                    //Trendlines
                    foreach (var tl in serie.TrendLines)
                    {
                        serieNo++;  
                        ApplyStyle(tl, Style.Trendline, serieNo);
                        if(tl.HasLbl) ApplyStyle(tl.Label, Style.TrendlineLabel, serieNo);
                    }

                    //Datapoints
                    if (serie is IDrawingChartDataPoints dps)
                    {
                        var items = serie.NumberOfItems;
                        foreach (var dp in dps.DataPoints)
                        {
                            applyBorder = !(chart.IsTypeStock() && dp.Border.Width == 0);
                            ApplyStyle(dp, dataPoint, dp.Index, items, applyFill, applyBorder);
                            if (dp.HasMarker())
                            {
                                ApplyStyle(dp.Marker, Style.DataPointMarker, dp.Index, items);
                                dp.Marker.Size = Style.DataPointMarkerLayout.Size;
                                if (Style.DataPointMarkerLayout.Style != eMarkerStyle.None)
                                {
                                    dp.Marker.Style = Style.DataPointMarkerLayout.Style;
                                }
                            }
                        }
                    }

                    //Errorbars
                    if(serie is ExcelChartSerieWithErrorBars chartSerieWithErrorBars && chartSerieWithErrorBars.ErrorBars!=null)
                    {
                        ApplyStyle(chartSerieWithErrorBars.ErrorBars, Style.ErrorBar);
                    }
                    serieNo++;
                }
            }
        }

        internal ExcelChartStyleEntry GetDataPointStyle(ExcelChart chart)
        {
            ExcelChartStyleEntry dataPoint;
            if (chart.IsType3D())
            {
                dataPoint = Style.DataPoint3D;
            }
            else if (chart.IsTypeLine() ||
                   (chart.IsTypeScatter() && chart.ChartType != eChartType.XYScatter) ||
                   (chart.IsTypeRadar() && chart.ChartType != eChartType.RadarFilled))
            {
                dataPoint = Style.DataPointLine;
            }
            else
            {
                dataPoint = Style.DataPoint;
            }

            return dataPoint;
        }

        internal void ApplyStyle(IDrawingStyleBase chartPart, ExcelChartStyleEntry section, int indexForColor=0, int numberOfItems=0, bool applyFill=true, bool applyBorder=true)
        {
            if(chartPart is IStyleMandatoryProperties setMandatoryProperties)
            {
                setMandatoryProperties.SetMandatoryProperties();
            }
            chartPart.CreatespPr();
            if(applyFill) ApplyStyleFill(chartPart, section, indexForColor, numberOfItems);

            if(applyBorder) ApplyStyleBorder(chartPart.Border, section, indexForColor, numberOfItems);
            ApplyStyleEffect(chartPart.Effect, section, indexForColor, numberOfItems);
            ApplyStyle3D(chartPart, section, indexForColor, numberOfItems);
            if (chartPart is IDrawingStyle chartPartWithFont)
            {
                ApplyStyleFont(section, indexForColor, chartPartWithFont, numberOfItems);
            }
        }
        private void ApplyStyleFill(IDrawingStyleBase chartPart, ExcelChartStyleEntry section, int indexForColor, int numberOfItems)
        {
            if (section.HasFill)  //Has inner fill section
            {
                chartPart.Fill.SetFromXml(section.Fill);
            }
            else if (section.FillReference.Index > 0)//From theme
            {
                var theme = GetTheme();
                if (theme.FormatScheme.FillStyle.Count > section.FillReference.Index - 1)
                {
                    var fill = theme.FormatScheme.FillStyle[section.FillReference.Index - 1];
                    chartPart.Fill.SetFromXml(fill);
                }
            }
            TransformColorFill(chartPart.Fill, section.FillReference.Color, indexForColor, numberOfItems);
        }

        private void ApplyStyleBorder(ExcelDrawingBorder chartBorder, ExcelChartStyleEntry section, int indexForColor,int numberOfItems)
        {
            if (section.HasBorder)  //Has border inner section
            {
                chartBorder.SetFromXml(section.Border.LineElement);
            }
            else if (section.BorderReference.Index > 0) //From theme
            {
                var theme = GetTheme();
                if (theme.FormatScheme.BorderStyle.Count > section.BorderReference.Index - 1)
                {
                    var border = theme.FormatScheme.BorderStyle[section.BorderReference.Index - 1];
                    chartBorder.SetFromXml(border.LineElement);
                }
            }
            TransformColorBorder(chartBorder, section.BorderReference.Color, indexForColor, numberOfItems);
        }

        private void ApplyStyleEffect(ExcelDrawingEffectStyle chartEffect, ExcelChartStyleEntry section, int indexForColor, int numberOfItems)
        {
            if (section.HasEffect)
            {
                chartEffect.SetFromXml(section.Effect.EffectElement);
            }
            else if (section.EffectReference.Index > 0) //From theme
            {
                var theme = GetTheme();
                if (theme.FormatScheme.EffectStyle.Count > section.EffectReference.Index - 1)
                {
                    var effect = theme.FormatScheme.EffectStyle[section.EffectReference.Index - 1];
                    chartEffect.SetFromXml(effect.Effect.EffectElement);
                }
            }
            TransformColorEffect(chartEffect, section.EffectReference.Color, indexForColor, numberOfItems);
        }
        private void ApplyStyle3D(IDrawingStyleBase chartThreeD, ExcelChartStyleEntry section, int indexForColor, int numberOfItems)
        {
            bool tranformColor = false;
            if (section.HasThreeD)
            {
                chartThreeD.ThreeD.SetFromXml(section.ThreeD.Scene3DElement, section.ThreeD.Sp3DElement);
                tranformColor = true;
            }
            else if (section.EffectReference.Index > 0) //From theme
            {
                var theme = GetTheme();
                if (theme.FormatScheme.EffectStyle.Count > section.EffectReference.Index - 1)
                {
                    var effect = theme.FormatScheme.EffectStyle[section.EffectReference.Index - 1];
                    chartThreeD.ThreeD.SetFromXml(effect.ThreeD.Scene3DElement, effect.ThreeD.Sp3DElement);
                    tranformColor = (effect.ThreeD.Sp3DElement != null);
                }
            }
            if (tranformColor)
            {
                TransformColorThreeD(chartThreeD.ThreeD, section.EffectReference.Color, indexForColor, numberOfItems);
            }
        }
        private void ApplyStyleFont(ExcelChartStyleEntry section, int indexForColor, IDrawingStyle chartPartWithFont, int numberOfItems)
        {
            if (section.HasTextBody)
            {
                chartPartWithFont.TextBody.SetFromXml(section.DefaultTextBody.PathElement);
            }
            if (section.HasTextRun)
            {                
                chartPartWithFont.Font.SetFromXml(section.DefaultTextRun.PathElement);                
            }

            if (section.FontReference.HasColor)
            {
                chartPartWithFont.Font.Fill.Style = eFillStyle.SolidFill;
                if (section.FontReference.Color.ColorType == eDrawingColorType.ChartStyleColor)
                {
                    ColorsManager.Transform(section.FontReference.Color, indexForColor == -1 ? 0 : indexForColor, numberOfItems);
                }
                chartPartWithFont.Font.Fill.SolidFill.Color.ApplyNewColor(section.FontReference.Color);
            }
            if (section.FontReference.Index != eThemeFontCollectionType.None)
            {
                chartPartWithFont.Font.LatinFont = $"+{(section.FontReference.Index == eThemeFontCollectionType.Minor ? "mn" : "mj")}-lt";
                chartPartWithFont.Font.EastAsianFont = $"+{(section.FontReference.Index == eThemeFontCollectionType.Minor ? "mn" : "mj")}-ea";
                chartPartWithFont.Font.ComplexFont = $"+{(section.FontReference.Index == eThemeFontCollectionType.Minor ? "mn" : "mj")}-cs";
            }
        }
        private void TransformColorBorder(ExcelDrawingBorder border, ExcelChartStyleColorManager color, int colorIndex, int numberOfItems)
        {
            TransformColorFillBasic(border.Fill, color, colorIndex, numberOfItems);
        }
        private void TransformColorFill(ExcelDrawingFill fill, ExcelChartStyleColorManager color, int colorIndex, int numberOfItems)
        {
            switch (fill.Style)
            {
                case eFillStyle.PatternFill:
                    TransformColor(fill.PatternFill.BackgroundColor, color, colorIndex, numberOfItems);
                    TransformColor(fill.PatternFill.ForegroundColor, color, colorIndex, numberOfItems);
                    break;
                case eFillStyle.BlipFill:
                    if(fill.BlipFill.Effects.Duotone!=null)
                    {
                        TransformColor(fill.BlipFill.Effects.Duotone.Color1, color, colorIndex, numberOfItems);
                        TransformColor(fill.BlipFill.Effects.Duotone.Color2, color, colorIndex, numberOfItems);
                    }
                    break;
                default:
                    TransformColorFillBasic(fill, color, colorIndex, numberOfItems);
                    break;
            }
        }

        private void TransformColorFillBasic(ExcelDrawingFillBasic fill, ExcelChartStyleColorManager color, int colorIndex, int numberOfItems)
        {
            switch (fill.Style)
            {
                case eFillStyle.SolidFill:
                    TransformColor(fill.SolidFill.Color, color, colorIndex, numberOfItems);
                    break;
                case eFillStyle.GradientFill:
                    foreach (var grad in fill.GradientFill.Colors)
                    {
                        TransformColor(grad.Color, color, colorIndex, numberOfItems);
                    }
                    break;
            }
        }
        private void TransformColorEffect(ExcelDrawingEffectStyle effect, ExcelChartStyleColorManager color, int colorIndex, int numberOfItems)
        {
            if (effect.HasInnerShadow && effect.InnerShadow.Color.ColorType == eDrawingColorType.Scheme && effect.InnerShadow.Color.SchemeColor.Color==eSchemeColor.Style)
            {
                TransformColor(effect.InnerShadow.Color, color, colorIndex, numberOfItems);
            }
            if(effect.HasOuterShadow && effect.OuterShadow.Color.ColorType == eDrawingColorType.Scheme && effect.OuterShadow.Color.SchemeColor.Color == eSchemeColor.Style)
            {
                TransformColor(effect.OuterShadow.Color, color, colorIndex, numberOfItems);
            }
            if (effect.HasPresetShadow && effect.PresetShadow.Color.ColorType == eDrawingColorType.Scheme && effect.PresetShadow.Color.SchemeColor.Color == eSchemeColor.Style)
            {
                TransformColor(effect.PresetShadow.Color, color, colorIndex, numberOfItems);
            }
            if (effect.HasGlow && effect.Glow.Color.ColorType == eDrawingColorType.Scheme && effect.Glow.Color.SchemeColor.Color == eSchemeColor.Style)
            {
                TransformColor(effect.Glow.Color, color, colorIndex, numberOfItems);
            }
            if (effect.HasFillOverlay)
            {
                TransformColorFill(effect.FillOverlay.Fill, color, colorIndex, numberOfItems);
            }
        }
        private void TransformColorThreeD(ExcelDrawing3D threeD, ExcelChartStyleColorManager color, int colorIndex, int numberOfItems)
        {
            if (threeD.ExtrusionColor.ColorType==eDrawingColorType.Scheme && threeD.ExtrusionColor.SchemeColor.Color == eSchemeColor.Style)
            {
                TransformColor(threeD.ExtrusionColor, color, colorIndex, numberOfItems);
            }
            if(threeD.ContourColor.ColorType == eDrawingColorType.Scheme && threeD.ContourColor.SchemeColor.Color == eSchemeColor.Style)
            {
                TransformColor(threeD.ContourColor, color, colorIndex, numberOfItems);
            }
        }
        private void TransformColor(ExcelDrawingColorManager color, ExcelChartStyleColorManager templateColor, int colorIndex, int numberOfItems)
        {
            if(templateColor!=null && templateColor.ColorType==eDrawingColorType.ChartStyleColor && color.ColorType == eDrawingColorType.Scheme && color.SchemeColor.Color == eSchemeColor.Style)
            {
                ColorsManager.Transform(color, templateColor.StyleColor.Index ?? colorIndex, numberOfItems);
            }
            else if (color.ColorType==eDrawingColorType.Scheme && color.SchemeColor.Color==eSchemeColor.Style)
            {
                ColorsManager.Transform(color, colorIndex, numberOfItems);
            }
        }

        private int CreateStylePart(ZipPackage p)
        {
            var id = GetIxFromChartUri(_chart.UriChart.OriginalString);
            StyleUri = GetNewUri(p, "/xl/charts/style{0}.xml", ref id);
            _chart.Part.CreateRelationship(StyleUri, TargetMode.Internal, ExcelPackage.schemaChartStyleRelationships);
            StylePart = p.CreatePart(StyleUri, ContentTypes.contentTypeChartStyle);
            return id;
        }
        private int CreateColorPart(ZipPackage p)
        {
            var id = GetIxFromChartUri(_chart.UriChart.OriginalString);
            ColorsUri = GetNewUri(p, "/xl/charts/colors{0}.xml", ref id);
            _chart.Part.CreateRelationship(ColorsUri, TargetMode.Internal, ExcelPackage.schemaChartColorStyleRelationships);
            ColorsPart = p.CreatePart(ColorsUri, ContentTypes.contentTypeChartColorStyle);
            return id;
        }
        private ZipPackageRelationship CreateThemeOverridePart(ZipPackage p, ZipPackagePart partToCopy)
        {
            var id = GetIxFromChartUri(_chart.UriChart.OriginalString);
            ThemeOverrideUri = GetNewUri(p, "/xl/theme/themeOverride{0}.xml", ref id);
            var rel=_chart.Part.CreateRelationship(ThemeOverrideUri, TargetMode.Internal, ExcelPackage.schemaThemeOverrideRelationships);
            ThemeOverridePart = p.CreatePart(ThemeOverrideUri, ContentTypes.contentTypeThemeOverride);

            ThemeOverrideXml = new XmlDocument();
            ThemeOverrideXml.Load(partToCopy.GetStream());

            foreach (var themeRel in partToCopy.GetRelationships())
            {
                var uri = OfficeOpenXml.Utils.UriHelper.ResolvePartUri(new Uri("/xl/chart/theme1.xml", UriKind.Relative), themeRel.TargetUri);
                var toPart = _chart.Part.Package.CreatePart(uri, PictureStore.GetContentType(uri.OriginalString));                
                var imageRel = ThemeOverridePart.CreateRelationship(uri, TargetMode.Internal, themeRel.RelationshipType);
                SetRelIdInThemeDoc(ThemeOverrideXml, themeRel.Id, imageRel.Id);
                var stream = partToCopy.GetStream();
                var b = ((MemoryStream)stream).GetBuffer();
                toPart.GetStream().Write(b, 0, b.Length);
            }

            ThemeOverrideXml.Save(ThemeOverridePart.GetStream(FileMode.CreateNew));
            partToCopy.Package.Dispose();
            return rel;
        }

        private void SetRelIdInThemeDoc(XmlDocument themeOverrideXml, string fromRelId, string toRelId)
        {
            foreach(XmlElement fill in themeOverrideXml.SelectNodes("//a:blipFill/a:blip", NameSpaceManager))
            { 
                
                if(fill!=null)
                {
                    var relId=fill.GetAttribute("r:embed");
                    if(relId==fromRelId)
                    {
                        fill.SetAttribute("r:embed", toRelId);
                    }
                }
            }
        }

        private string GetStartStyleXml(int id)
        {
            var sb = new StringBuilder();
            sb.Append($"<cs:chartStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" id=\"{id}\">");
            AppendDefaultStyleSection(sb, "axisTitle");
            AppendDefaultStyleSection(sb, "categoryAxis");
            AppendDefaultStyleSection(sb, "chartArea");
            AppendDefaultStyleSection(sb, "dataLabel");
            AppendDefaultStyleSection(sb, "dataLabelCallout");
            AppendDefaultStyleSection(sb, "dataPoint");
            AppendDefaultStyleSection(sb, "dataPoint3D");
            AppendDefaultStyleSection(sb, "dataPointLine");
            AppendDefaultStyleSection(sb, "dataPointMarker");
            sb.Append("<cs:dataPointMarkerLayout size=\"17\" symbol=\"circle\"/>");
            AppendDefaultStyleSection(sb, "dataPointWireframe");
            AppendDefaultStyleSection(sb, "dataTable");
            AppendDefaultStyleSection(sb, "downBar");
            AppendDefaultStyleSection(sb, "dropLine");
            AppendDefaultStyleSection(sb, "errorBar");
            AppendDefaultStyleSection(sb, "floor");
            AppendDefaultStyleSection(sb, "gridlineMajor");
            AppendDefaultStyleSection(sb, "gridlineMinor");
            AppendDefaultStyleSection(sb, "hiLoLine");
            AppendDefaultStyleSection(sb, "leaderLine");
            AppendDefaultStyleSection(sb, "legend");
            AppendDefaultStyleSection(sb, "plotArea");
            AppendDefaultStyleSection(sb, "plotArea3D");
            AppendDefaultStyleSection(sb, "seriesAxis");
            AppendDefaultStyleSection(sb, "seriesLine");
            AppendDefaultStyleSection(sb, "title");
            AppendDefaultStyleSection(sb, "trendline");
            AppendDefaultStyleSection(sb, "trendlineLabel");
            AppendDefaultStyleSection(sb, "upBar");
            AppendDefaultStyleSection(sb, "valueAxis");
            AppendDefaultStyleSection(sb, "wall");
            sb.Append($"</cs:chartStyle>");

            return sb.ToString();
        }
        private void AppendDefaultStyleSection(StringBuilder sb, string section)
        {
            sb.Append($"<cs:{section}><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"></cs:fontRef></cs:{section}>");
        }
        private string GetStartColorXml()
        {
            return $"<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" meth=\"cycle\" id=\"10\">" +
                "<a:schemeClr val=\"accent1\"/>" +
                "<a:schemeClr val=\"accent2\"/>" +
                "<a:schemeClr val=\"accent3\"/>" +
                "<a:schemeClr val=\"accent4\"/>" +
                "<a:schemeClr val=\"accent5\"/>" +
                "<a:schemeClr val=\"accent6\"/>" +
                "<cs:variation/><cs:variation><a:lumMod val=\"60000\"/></cs:variation>" +
                "<cs:variation><a:lumMod val=\"80000\"/><a:lumOff val=\"20000\"/></cs:variation>" +
                "<cs:variation><a:lumMod val=\"80000\"/></cs:variation>" +
                "<cs:variation><a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/></cs:variation>" +
                "<cs:variation><a:lumMod val=\"50000\"/></cs:variation>" +
                "<cs:variation><a:lumMod val=\"70000\"/><a:lumOff val=\"30000\"/></cs:variation>" +
                "<cs:variation><a:lumMod val=\"70000\"/></cs:variation>" +
                "<cs:variation><a:lumMod val=\"50000\"/><a:lumOff val=\"50000\"/></cs:variation>" +
                "</cs:colorStyle>";
        }
        private int GetIxFromChartUri(string name)
        {
            if (name.StartsWith("chart", StringComparison.InvariantCultureIgnoreCase) && name.EndsWith(".xml", StringComparison.InvariantCultureIgnoreCase))
            {
                var n = name.Substring(5, name.Length - 9);
                try
                {
                    return int.Parse(n);
                }
                catch
                {
                    return 1;
                }

            }
            return 1;
        }
        private void LoadStyleAndColors(ExcelChart chart)
        {
            if (chart.Part == null) return;
            var p = chart.WorkSheet.Workbook._package;
            foreach (var rel in chart.Part.GetRelationships())
            {
                if (rel.RelationshipType == ExcelPackage.schemaChartStyleRelationships)
                {
                    StyleUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    StylePart = p.ZipPackage.GetPart(StyleUri);
                    StyleXml = new XmlDocument();
                    LoadXmlSafe(StyleXml, StylePart.GetStream());
                }
                else if (rel.RelationshipType == ExcelPackage.schemaChartColorStyleRelationships)
                {
                    ColorsUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    ColorsPart = p.ZipPackage.GetPart(ColorsUri);
                    ColorsXml = new XmlDocument();
                    LoadXmlSafe(ColorsXml, ColorsPart.GetStream());
                }
            }
        }
        private ExcelThemeBase GetTheme()
        {
            if (ThemeOverride == null)
            {
                return _chart.WorkSheet.Workbook.ThemeManager.CurrentTheme;
            }
            else
            {
                return ThemeOverride;
            }
        }

        /// <summary>
        /// A reference to style settings for the chart
        /// </summary>
        public ExcelChartStyle Style
        {
            get;
            private set;
        }
        /// <summary>
        /// A reference to color style settings for the chart
        /// </summary>
        public ExcelChartColorsManager ColorsManager
        {
            get;
            private set;
        }
        /// <summary>
        /// If the chart has a different theme than the theme in the workbook, this property defines that theme.
        /// </summary>
        public ExcelThemeOverride ThemeOverride
        {
            get;
            private set;
        } = null;
        /// <summary>
        /// The chart style xml document
        /// </summary>
        public XmlDocument StyleXml { get; private set; }
        /// <summary>
        /// The color xml document
        /// </summary>
        public XmlDocument ColorsXml { get; private set; }
        /// <summary>
        /// Overrides the current theme for the chart.
        /// </summary>
        public XmlDocument ThemeOverrideXml { get; private set; }
        internal Uri StyleUri { get; set; }
        internal ZipPackagePart StylePart { get; set; }
        internal Uri ColorsUri { get; set; }
        internal ZipPackagePart ColorsPart { get; set; }
        internal Uri ThemeOverrideUri { get; set; }
        internal ZipPackagePart ThemeOverridePart { get; set; }
    }
}
