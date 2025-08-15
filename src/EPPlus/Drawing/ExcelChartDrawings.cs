using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Collection of drawings inside a chart
    /// </summary>
    public class ExcelChartDrawings : IEnumerable<ExcelDrawing>, IDisposable//, IPictureRelationDocument
    {
        internal ExcelDrawings _drawings = null;
        internal ExcelChart _chart;

        internal XmlNamespaceManager NameSpaceManager { get{ return _drawings.NameSpaceManager; } }

        internal Packaging.ZipPackagePart Part { get { return _drawings.Part; } }

        internal int _nextDrawingId { get { return _drawings._nextDrawingId; } set { _drawings._nextDrawingId = value; } }

        /// <summary>
        /// A reference to the drawing xml document
        /// </summary>
        public XmlDocument DrawingXml
        {
            get
            {
                return _drawings.DrawingXml;
            }
        }

        internal ExcelChartDrawings(ExcelChartStandard chart)
        {
            _chart = chart;
            LoadDrawings(chart);
        }

        internal void LoadDrawings(ExcelChartStandard chart)
        {
            if (_drawings == null)
            {
                _drawings = new ExcelDrawings(_chart.WorkSheet._package, _chart);
            }
        }

        /// <summary>
        /// Adds a shape to the chart
        /// </summary>
        /// <param name="Name">The name of the shape</param>
        /// <param name="Style">The type of shape</param>
        /// <returns>The shape</returns>
        public ExcelShape AddShape(string Name, eShapeStyle Style)
        {
            return _drawings.AddShape(Name, Style, _chart);
        }

        /// <summary>
        /// Adds a picture to the chart
        /// </summary>
        /// <param name="Name">The name of the shape</param>
        /// <param name="ImagePath">The path to the image file.</param>
        /// <returns>The shape</returns>
        public ExcelPicture AddPicture(string Name, string ImagePath)
        {
            return AddPicture(Name, new FileInfo(ImagePath), null);
        }
        /// <summary>
        /// Adds a picture to the chart
        /// </summary>
        /// <param name="Name">The name of the shape</param>
        /// <param name="ImagePath">The path to the image file.</param>
        /// <param name="HyperLink">A hyperlink for the shape</param>
        /// <returns>The shape</returns>
        public ExcelPicture AddPicture(string Name, string ImagePath, ExcelHyperLink HyperLink)
        {
            return AddPicture(Name, new FileInfo(ImagePath), HyperLink);
        }

        /// <summary>
        /// Adds a picture to the chart
        /// </summary>
        /// <param name="Name">The name of the shape</param>
        /// <param name="ImageFile">The image file.</param>
        /// <returns>The shape</returns>
        public ExcelPicture AddPicture(string Name, FileInfo ImageFile)
        {
            return AddPicture(Name, ImageFile, null);
        }
        /// <summary>
        /// Adds a picture to the chart
        /// </summary>
        /// <param name="Name">The name of the shape</param>
        /// <param name="ImageFile">The image file.</param>
        /// <param name="HyperLink">A hyperlink for the shape</param>
        /// <returns>The shape</returns>
        public ExcelPicture AddPicture(string Name, FileInfo ImageFile, Uri HyperLink)
        {
            return _drawings.AddPicture(Name, ImageFile, HyperLink, PictureLocation.Embed, _chart);
        }

        /// <summary>
        /// Adds a picture to the chart
        /// </summary>
        /// <param name="Name">The name of the shape</param>
        /// <param name="ImageStream">The stream containing image file.</param>
        /// <returns>The shape</returns>
        public ExcelPicture AddPicture(string Name, Stream ImageStream)
        {
            return AddPicture(Name, ImageStream, null);
        }
        /// <summary>
        /// Adds a picture to the chart
        /// </summary>
        /// <param name="Name">The name of the shape</param>
        /// <param name="ImageStream">The stream containing image file.</param>
        /// <param name="HyperLink">A hyperlink for the shape</param>
        /// <returns>The shape</returns>
        public ExcelPicture AddPicture(string Name, Stream ImageStream, Uri HyperLink)
        {
            return _drawings.AddPicture(Name, ImageStream, HyperLink, _chart);
        }

        internal void AddDrawingInternal(ExcelDrawing dr)
        {
            _drawings.AddDrawingInternal(dr);
        }

        internal XmlElement CreateDocumentAndTopNodeChartDrawings(ExcelChart ContainerChart)
        {
            return _drawings.CreateDocumentAndTopNodeChartDrawings(ContainerChart);
        }

        internal XmlElement CreateDrawingXmlChartDrawings(ExcelChart container)
        {
            return _drawings.CreateDrawingXmlChartDrawings(container);
        }

        internal string GetUniqueDrawingName(string name)
        {
            return _drawings.GetUniqueDrawingName(name);
        }


        /// <summary>
        /// Returns the drawing at the specified position.  
        /// </summary>
        /// <param name="PositionID">The position of the drawing. 0-base</param>
        /// <returns>The drawing</returns>
        public ExcelDrawing this[int PositionID]
        {
            get
            {
                return (_drawings._drawingsList[PositionID]);
            }
        }
        /// <summary>
        /// Returns the drawing matching the specified name
        /// </summary>
        /// <param name="Name">The name of the drawing</param>
        /// <returns>The drawing</returns>
        public ExcelDrawing this[string Name]
        {
            get
            {
                if (_drawings._drawingNames.ContainsKey(Name))
                {
                    return _drawings._drawingsList[_drawings._drawingNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _drawings.Count;
            }
        }

        /// <summary>
        /// Removes a drawing.
        /// </summary>
        /// <param name="Index">The index of the drawing</param>
        public void Remove(int Index)
        {
            if (_drawings.Worksheet is ExcelChartsheet && _drawings._drawingsList.Count > 0)
            {
                throw new InvalidOperationException("Can't remove charts from chart worksheets");
            }
            _drawings.RemoveDrawing(Index);
        }
        /// <summary>
        /// Removes a drawing.
        /// </summary>
        /// <param name="Drawing">The drawing</param>
        public void Remove(ExcelDrawing Drawing)
        {
            _drawings.Remove(_drawings._drawingNames[Drawing.Name]);
        }
        /// <summary>
        /// Removes a drawing.
        /// </summary>
        /// <param name="Name">The name of the drawing</param>
        public void Remove(string Name)
        {
            _drawings.Remove(_drawings._drawingNames[Name]);
        }

        /// <summary>
        /// Returns an Enumerator for the drawings in the chart
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelDrawing> GetEnumerator()
        {
            return _drawings._drawingsList.GetEnumerator();
        }
        /// <summary>
        /// Returns an Enumerator for the drawings in the chart
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
        /// <summary>
        /// Disposes the drawings collection and all drawings in it.
        /// </summary>
        public void Dispose()
        {
            if (_drawings != null)
            {
                foreach (var d in _drawings)
                {
                    d.Dispose();
                }
            }
        }
    }
}
