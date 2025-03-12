using OfficeOpenXml.Drawing.Chart;
using System;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelChartDrawings
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

        internal ExcelChartDrawings(ExcelChart chart)
        {
            _chart = chart;
            LoadDrawings(chart);
        }

        internal void LoadDrawings(ExcelChart chart)
        {
            if (_drawings == null)
            {
                _drawings = new ExcelDrawings(_chart.WorkSheet._package, _chart);
            }
        }

        public ExcelShape AddShape(string Name, eShapeStyle Style)
        {
            return _drawings.AddShape(Name, Style, _chart);
        }

        //String

        public ExcelPicture AddPicture(string Name, string ImagePath)
        {
            return AddPicture(Name, new FileInfo(ImagePath), null);
        }
        public ExcelPicture AddPicture(string Name, string ImagePath, ExcelHyperLink HyperLink)
        {
            return AddPicture(Name, new FileInfo(ImagePath), HyperLink);
        }

        //FileInfo

        public ExcelPicture AddPicture(string Name, FileInfo ImagePath)
        {
            return AddPicture(Name, ImagePath, null);
        }
        public ExcelPicture AddPicture(string Name, FileInfo ImagePath, Uri HyperLink)
        {
            return _drawings.AddPicture(Name, ImagePath, HyperLink, PictureLocation.Embed, _chart);
        }

        //Stream

        public ExcelPicture AddPicture(string Name, Stream ImagePath)
        {
            return AddPicture(Name, ImagePath, null);
        }
        public ExcelPicture AddPicture(string Name, Stream ImagePath, Uri HyperLink)
        {
            return _drawings.AddPicture(Name, ImagePath, HyperLink, _chart);
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
        /// <returns></returns>
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
        /// <param name="Name">The name of the worksheet</param>
        /// <returns></returns>
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
    }
}
