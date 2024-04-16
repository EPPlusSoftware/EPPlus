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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// The state of the pane.
    /// </summary>
    public enum ePaneState
    {
        /// <summary>
        /// Panes are frozen, but were not split being frozen.In this state, when the panes are unfrozen again, a single pane results, with no split. In this state, the split bars are not adjustable.
        /// </summary>
        Frozen,
        /// <summary>
        /// Frozen Split
        /// Panes are frozen and were split before being frozen. In this state, when the panes are unfrozen again, the split remains, but is adjustable.
        /// </summary>
        FrozenSplit,
        /// <summary>
        /// Panes are split, but not frozen.In this state, the split bars are adjustable by the user.
        /// </summary>
        Split
    }
    /// <summary>
    /// The position of the pane.
    /// </summary>
    public enum ePanePosition
    {
        /// <summary>
        /// Bottom Left Pane.
        /// Used when worksheet view has both vertical and horizontal splits.
        /// Also used when the worksheet is horizontaly split only, specifying this is the bottom pane.
        /// </summary>
        BottomLeft,
        /// <summary>
        /// Bottom Right Pane. 
        /// This property is only used when the worksheet has both vertical and horizontal splits.
        /// </summary>
        BottomRight,
        /// <summary>
        /// Top Left Pane.
        /// Used when worksheet view has both vertical and horizontal splits.
        /// Also used when the worksheet is horizontaly split only, specifying this is the top pane.
        /// </summary>
        TopLeft,
        /// <summary>
        /// Top Right Pane
        /// Used when the worksheet view has both vertical and horizontal splits.
        /// Also used when the worksheet is verticaly split only, specifying this is the right pane.
        /// </summary>
        TopRight
    }

    /// <summary>
    /// Represents the different view states of the worksheet
    /// </summary>
    public class ExcelWorksheetView : XmlHelper
    {
        /// <summary>
        /// Defines general properties for the panes, if the worksheet is frozen or split.
        /// </summary>
        public class ExcelWorksheetViewPaneSettings : XmlHelper
        {
            internal ExcelWorksheetViewPaneSettings(XmlNamespaceManager ns, XmlNode topNode) :
                base(ns, topNode)
            {
            }
            /// <summary>
            /// The state of the pane.
            /// </summary>
            public ePaneState State
            {
                get
                {
                    return GetXmlEnumNull<ePaneState>("@state", ePaneState.Split).Value;
                }
                internal set
                {
                    SetXmlNodeString("@state", value.ToEnumString());
                }
            }
            /// <summary>
            /// The active pane
            /// </summary>
            public ePanePosition ActivePanePosition
            {
                get
                {
                    return GetXmlEnumNull<ePanePosition>("@activePane", ePanePosition.TopLeft).Value;
                }
                set
                {
                    SetXmlNodeString("@activePane", value.ToEnumString());
                }
            }

            /// <summary>
            /// The horizontal position of the split. 1/20 of a point if the pane is split. Number of columns in the top pane if this pane is frozen.
            /// </summary>
            public double XSplit
            {
                get
                {
                    return GetXmlNodeDouble("@xSplit");
                }
                set
                {
                    SetXmlNodeDouble("@xSplit", value, false);
                }
            }
            /// <summary>
            /// The vertical position of the split. 1/20 of a point if the pane is split. Number of rows in the left pane if this pane is frozen.
            /// </summary>
            public double YSplit
            {
                get
                {
                    return GetXmlNodeDouble("@ySplit");
                }
                set
                {
                    SetXmlNodeDouble("@ySplit", value, false);
                }
            }
            /// <summary>
            /// 
            /// </summary>
            public string TopLeftCell
            {
                get
                {
                    return GetXmlNodeString("@topLeftCell");
                }
                set
                {
                    if (string.IsNullOrEmpty(value))
                    {
                        DeleteNode("@topLeftCell");
                    }
                    else if (ExcelCellBase.IsValidCellAddress(value))
                    {
                        SetXmlNodeString("@topLeftCell", value);
                    }
                    else
                    {
                        throw new InvalidOperationException("The value must be a value cell address");
                    }
                }
            }
            internal static XmlNode CreatePaneElement(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
            {
                var node = topNode.SelectSingleNode("d:pane", nameSpaceManager);
                if (node == null)
                {
                    node = topNode.OwnerDocument.CreateElement("pane", ExcelPackage.schemaMain);
                    topNode.PrependChild(node);
                }
                return node;
            }
        }
        /// <summary>
        /// The selection properties for panes after a freeze or split.
        /// </summary>
        public class ExcelWorksheetPanes : XmlHelper
        {
            XmlElement _selectionNode = null;
            internal ExcelWorksheetPanes(XmlNamespaceManager ns, XmlNode topNode) :
                base(ns, topNode)
            {
                if (topNode.Name == "selection")
                {
                    _selectionNode = topNode as XmlElement;
                }
            }

            const string _activeCellPath = "@activeCell";
            /// <summary>
            /// Set the active cell. Must be set within the SelectedRange.
            /// </summary>
            public string ActiveCell
            {
                get
                {
                    string address = GetXmlNodeString(_activeCellPath);
                    if (address == "")
                    {
                        return "A1";
                    }
                    return address;
                }
                set
                {
                    int fromCol, fromRow, toCol, toRow;
                    if (_selectionNode == null) CreateSelectionElement();
                    ExcelCellBase.GetRowColFromAddress(value, out fromRow, out fromCol, out toRow, out toCol);
                    SetXmlNodeString(_activeCellPath, value);
                    if (((XmlElement)TopNode).GetAttribute("sqref") == "")
                    {

                        SelectedRange = ExcelCellBase.GetAddress(fromRow, fromCol);
                    }
                    else
                    {
                        //TODO:Add fix for out of range here
                    }
                }
            }
            /// <summary>
            /// The position of the pane.
            /// </summary>
            public ePanePosition Position
            {
                get
                {
                    return GetXmlEnumNull<ePanePosition>("@pane", ePanePosition.TopLeft).Value;
                }
            }
            /// <summary>
            /// 
            /// </summary>
            public int ActiveCellId
            {
                get;
                set;
            }
            private void CreateSelectionElement()
            {
                _selectionNode = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                TopNode.AppendChild(_selectionNode);
                TopNode = _selectionNode;
            }
            const string _selectionRangePath = "@sqref";
            /// <summary>
            /// Selected Cells. Used in combination with ActiveCell
            /// </summary>        
            public string SelectedRange
            {
                get
                {
                    string address = GetXmlNodeString(_selectionRangePath);
                    if (address == "")
                    {
                        return "A1";
                    }
                    return address;
                }
                set
                {
                    int fromCol, fromRow, toCol, toRow;
                    if (_selectionNode == null) CreateSelectionElement();
                    ExcelCellBase.GetRowColFromAddress(value, out fromRow, out fromCol, out toRow, out toCol);
                    SetXmlNodeString(_selectionRangePath, value);
                    if (((XmlElement)TopNode).GetAttribute("activeCell") == "")
                    {

                        ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);
                    }
                    else
                    {
                        //TODO:Add fix for out of range here
                    }
                }
            }

        }
        private ExcelWorksheet _worksheet;

        #region ExcelWorksheetView Constructor
        /// <summary>
        /// Creates a new ExcelWorksheetView which provides access to all the view states of the worksheet.
        /// </summary>
        /// <param name="ns"></param>
        /// <param name="node"></param>
        /// <param name="xlWorksheet"></param>
        internal ExcelWorksheetView(XmlNamespaceManager ns, XmlNode node, ExcelWorksheet xlWorksheet) :
            base(ns, node)
        {
            _worksheet = xlWorksheet;
            SchemaNodeOrder = new string[] { "sheetViews", "sheetView", "pane", "selection" };
            if (_paneSettings == null)
            {
                _paneSettings = new ExcelWorksheetViewPaneSettings(NameSpaceManager, TopNode);
            }
            SetPaneSettings();
            Panes = LoadPanes();
        }

        private void SetPaneSettings()
        {
            var n = GetNode("d:pane");
            if (n == null)
            {
                PaneSettings = null;
            }
            else
            {
                PaneSettings = new ExcelWorksheetViewPaneSettings(NameSpaceManager, n);
            }
        }

        #endregion
        private ExcelWorksheetPanes[] LoadPanes()
        {
            XmlNodeList nodes = TopNode.SelectNodes("//d:selection", NameSpaceManager);
            if (nodes.Count == 0)
            {
                return new ExcelWorksheetPanes[] { new ExcelWorksheetPanes(NameSpaceManager, TopNode) };
            }
            else
            {
                ExcelWorksheetPanes[] panes = new ExcelWorksheetPanes[nodes.Count];
                int i = 0;
                foreach (XmlElement elem in nodes)
                {
                    panes[i++] = new ExcelWorksheetPanes(NameSpaceManager, elem);
                }
                return panes;
            }
        }
        #region SheetViewElement
        /// <summary>
        /// Returns a reference to the sheetView element
        /// </summary>
        protected internal XmlElement SheetViewElement
        {
            get
            {
                return (XmlElement)TopNode;
            }
        }
        #endregion
        #region Public Methods & Properties
        /// <summary>
        /// The active cell. Single cell address.                
        /// This cell must be inside the selected range. If not, the selected range is set to the active cell address
        /// </summary>
        public string ActiveCell
        {
            get
            {
                return Panes[Panes.GetUpperBound(0)].ActiveCell;
            }
            set
            {
                var ac = new ExcelAddressBase(value);
                if (ac.IsSingleCell == false)
                {
                    throw (new InvalidOperationException("ActiveCell must be a single cell."));
                }

                /*** Active cell must be inside SelectedRange ***/
                var sd = new ExcelAddressBase(SelectedRange.Replace(" ", ","));
                Panes[Panes.GetUpperBound(0)].ActiveCell = value;

                if (IsActiveCellInSelection(ac, sd) == false)
                {
                    SelectedRange = value;
                }
            }
        }
        /// <summary>
        /// The Top-Left Cell visible. Single cell address.
        /// Empty string or null is the same as A1.
        /// </summary>
        public string TopLeftCell
        {
            get
            {
                return GetXmlNodeString("@topLeftCell");
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    DeleteNode("@topLeftCell");
                }
                else
                {
                    if (!ExcelAddressBase.IsValidCellAddress(value))
                    {
                        throw (new InvalidOperationException("Must be a valid cell address."));
                    }
                    var ac = new ExcelAddressBase(value);
                    if (ac.IsSingleCell == false)
                    {
                        throw (new InvalidOperationException("ActiveCell must be a single cell."));
                    }
                    SetXmlNodeString("@topLeftCell", value);
                }
            }
        }

        /// <summary>
        /// Selected Cells in the worksheet. Used in combination with ActiveCell.
        /// If the active cell is not inside the selected range, the active cell will be set to the first cell in the selected range.
        /// If the selected range has multiple adresses, these are separated with space. If the active cell is not within the first address in this list, the attribute ActiveCellId must be set (not supported, so it must be set via the XML).
        /// </summary>
        public string SelectedRange
        {
            get
            {
                return Panes[Panes.GetUpperBound(0)].SelectedRange;
            }
            set
            {
                var ac = new ExcelAddressBase(ActiveCell);

                /*** Active cell must be inside SelectedRange ***/
                var sd = new ExcelAddressBase(value.Replace(" ", ","));      //Space delimitered here, replace

                Panes[Panes.GetUpperBound(0)].SelectedRange = value;
                if (IsActiveCellInSelection(ac, sd) == false)
                {
                    ActiveCell = new ExcelCellAddress(sd._fromRow, sd._fromCol).Address;
                }
            }
        }
        ExcelWorksheetViewPaneSettings _paneSettings = null;
        /// <summary>
        /// Contains settings for the active pane
        /// </summary>
        public ExcelWorksheetViewPaneSettings PaneSettings
        {
            get;
            private set;
        }

        private bool IsActiveCellInSelection(ExcelAddressBase ac, ExcelAddressBase sd)
        {
            var c = sd.Collide(ac);
            if (c == ExcelAddressBase.eAddressCollition.Equal || c == ExcelAddressBase.eAddressCollition.Inside)
            {
                return true;
            }
            else
            {
                if (sd.Addresses != null)
                {
                    foreach (var sds in sd.Addresses)
                    {
                        c = sds.Collide(ac);
                        if (c == ExcelAddressBase.eAddressCollition.Equal || c == ExcelAddressBase.eAddressCollition.Inside)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// If the worksheet is selected within the workbook. NOTE: Setter clears other selected tabs.
        /// </summary>
        public bool TabSelected
        {
            get
            {
                return GetXmlNodeBool("@tabSelected");
            }
            set
            {
                SetTabSelected(value, false);
            }
        }

        /// <summary>
        /// If the worksheet is selected within the workbook. NOTE: Setter keeps other selected tabs.
        /// </summary>
        public bool TabSelectedMulti
        {
            get
            {
                return GetXmlNodeBool("@tabSelected");
            }
            set
            {
                SetTabSelected(value, true);
            }
        }

        /// <summary>
        /// Sets whether the worksheet is selected within the workbook.
        /// </summary>
        /// <param name="isSelected">Whether the tab is selected, defaults to true.</param>
        /// <param name="allowMultiple">Whether to allow multiple active tabs, defaults to false.</param>
        public void SetTabSelected(bool isSelected = true, bool allowMultiple = false)
        {
            if (isSelected)
            {
                SheetViewElement.SetAttribute("tabSelected", "1");
                if (!allowMultiple)
                {
                    //    // ensure no other worksheet has its tabSelected attribute set to 1
                    foreach (ExcelWorksheet sheet in _worksheet._package.Workbook.Worksheets)
                        sheet.View.TabSelected = false;

                }
                XmlElement bookView = _worksheet.Workbook.WorkbookXml.SelectSingleNode("//d:workbookView", _worksheet.NameSpaceManager) as XmlElement;
                if (bookView != null)
                {
                    bookView.SetAttribute("activeTab", (_worksheet.PositionId).ToString());
                }
            }
            else
                SetXmlNodeString("@tabSelected", "0");
        }

        /// <summary>
        /// Sets the view mode of the worksheet to pagelayout
        /// </summary>
        public bool PageLayoutView
        {
            get
            {
                return GetXmlNodeString("@view") == "pageLayout";
            }
            set
            {
                if (value)
                    SetXmlNodeString("@view", "pageLayout");
                else
                    SheetViewElement.RemoveAttribute("view");
            }
        }
        /// <summary>
        /// Sets the view mode of the worksheet to pagebreak
        /// </summary>
        public bool PageBreakView
        {
            get
            {
                return GetXmlNodeString("@view") == "pageBreakPreview";
            }
            set
            {
                if (value)
                    SetXmlNodeString("@view", "pageBreakPreview");
                else
                    SheetViewElement.RemoveAttribute("view");
            }
        }
        /// <summary>
        /// Show gridlines in the worksheet
        /// </summary>
        public bool ShowGridLines
        {
            get
            {
                return GetXmlNodeBool("@showGridLines", true);
            }
            set
            {
                SetXmlNodeString("@showGridLines", value ? "1" : "0");
            }
        }
        /// <summary>
        /// Show the Column/Row headers (containg column letters and row numbers)
        /// </summary>
        public bool ShowHeaders
        {
            get
            {
                return GetXmlNodeBool("@showRowColHeaders", true);
            }
            set
            {
                SetXmlNodeString("@showRowColHeaders", value ? "1" : "0");
            }
        }
        /// <summary>
        /// Window zoom magnification for current view representing percent values.
        /// </summary>
        public int ZoomScale
        {
            get
            {
                return GetXmlNodeInt("@zoomScale");
            }
            set
            {
                if (value < 10 || value > 400)
                {
                    throw new ArgumentOutOfRangeException("Zoome scale out of range (10-400)");
                }
                SetXmlNodeString("@zoomScale", value.ToString());
            }
        }
        /// <summary>
        /// If the sheet is in 'right to left' display mode. Column A is on the far right and column B to the left of A. Text is also 'right to left'.
        /// </summary>
        public bool RightToLeft
        {
            get
            {
                return GetXmlNodeBool("@rightToLeft");
            }
            set
            {
                SetXmlNodeString("@rightToLeft", value == true ? "1" : "0");
            }
        }
        internal bool WindowProtection
        {
            get
            {
                return GetXmlNodeBool("@windowProtection", false);
            }
            set
            {
                SetXmlNodeBool("@windowProtection", value, false);
            }
        }
        /// <summary>
        /// Reference to the panes
        /// </summary>
        public ExcelWorksheetPanes[] Panes
        {
            get;
            internal set;
        }
        /// <summary>
        /// The top left pane or the top pane if the sheet is horizontaly split. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
        /// </summary>
        public ExcelWorksheetPanes TopLeftPane
        {
            get
            {
                return Panes?.Where(x => x.Position == ePanePosition.TopLeft).FirstOrDefault();
            }
        }
        /// <summary>
        /// The top right pane. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
        /// </summary>
        public ExcelWorksheetPanes TopRightPane
        {
            get
            {
                return Panes?.Where(x => x.Position == ePanePosition.TopRight).FirstOrDefault();
            }
        }
        /// <summary>
        /// The bottom left pane. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
        /// </summary>
        public ExcelWorksheetPanes BottomLeftPane
        {
            get
            {
                return Panes?.Where(x => x.Position == ePanePosition.BottomLeft).FirstOrDefault();
            }
        }
        /// <summary>
        /// The bottom right pane. This property returns null if the pane does not exist in the <see cref="Panes"/> array.
        /// </summary>
        public ExcelWorksheetPanes BottomRightPane
        {
            get
            {
                return Panes?.Where(x => x.Position == ePanePosition.BottomRight).FirstOrDefault();
            }
        }
        string _paneNodePath = "d:pane";
        string _selectionNodePath = "d:selection";
        /// <summary>
        /// Freeze the columns/rows to left and above the cell
        /// </summary>
        /// <param name="Row"></param>
        /// <param name="Column"></param>
        public void FreezePanes(int Row, int Column)
        {
            //TODO:fix this method to handle splits as well.
            ValidateRows(Row, Column);

            if (Row == 1 && Column == 1)
            {
                UnFreezePanes();
                return;
            }

            bool isSplit;
            if (PaneSettings == null)
            {
                var node = ExcelWorksheetViewPaneSettings.CreatePaneElement(NameSpaceManager, TopNode);
                PaneSettings = new ExcelWorksheetViewPaneSettings(NameSpaceManager, node);
                isSplit = false;
            }
            else
            {

                isSplit = PaneSettings.State != ePaneState.Frozen;
                PaneSettings.TopNode.RemoveAll();
            }

            if (Column > 1) PaneSettings.XSplit = Column - 1;
            if (Row > 1) PaneSettings.YSplit = Row - 1;
            PaneSettings.TopLeftCell = ExcelCellBase.GetAddress(Row, Column);
            PaneSettings.State = isSplit ? ePaneState.FrozenSplit : ePaneState.Frozen;

            CreateSelectionXml(Row - 1, Column - 1, false);
            Panes = LoadPanes();
        }

        private void CreateSelectionXml(int Row, int Column, bool isSplit)
        {
            RemoveSelection();

            string sqRef = SelectedRange, activeCell = ActiveCell;
            PaneSettings.ActivePanePosition = ePanePosition.BottomRight;
            XmlNode afterNode;
            if (isSplit)
            {
                //Top left node, default pane
                afterNode = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                PaneSettings.TopNode.ParentNode.InsertAfter(afterNode, PaneSettings.TopNode);
            }
            else
            {
                afterNode = PaneSettings.TopNode;
            }

            if (Row > 0 && Column == 0)
            {
                PaneSettings.ActivePanePosition = ePanePosition.BottomLeft;
                XmlElement sel = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                sel.SetAttribute("pane", "bottomLeft");
                if (activeCell != "") sel.SetAttribute("activeCell", activeCell);
                if (sqRef != "") sel.SetAttribute("sqref", sqRef);
                TopNode.InsertAfter(sel, afterNode);
            }
            else if (Column > 0 && Row == 0)
            {
                PaneSettings.ActivePanePosition = ePanePosition.TopRight;
                XmlElement sel = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                sel.SetAttribute("pane", "topRight");
                if (activeCell != "") sel.SetAttribute("activeCell", activeCell);
                if (sqRef != "") sel.SetAttribute("sqref", sqRef);
                TopNode.InsertAfter(sel, afterNode);
            }
            else
            {
                XmlElement selTR = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                selTR.SetAttribute("pane", "topRight");
                string cell = ExcelCellBase.GetAddress(1, Column + 1);
                selTR.SetAttribute("activeCell", cell);
                selTR.SetAttribute("sqref", cell);
                afterNode.ParentNode.InsertAfter(selTR, afterNode);

                XmlElement selBL = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                cell = ExcelCellBase.GetAddress(Row + 1, 1);
                selBL.SetAttribute("pane", "bottomLeft");
                selBL.SetAttribute("activeCell", cell);
                selBL.SetAttribute("sqref", cell);
                selTR.ParentNode.InsertAfter(selBL, selTR);


                XmlElement selBR = TopNode.OwnerDocument.CreateElement("selection", ExcelPackage.schemaMain);
                selBR.SetAttribute("pane", "bottomRight");
                if (activeCell != "") selBR.SetAttribute("activeCell", activeCell);
                if (sqRef != "") selBR.SetAttribute("sqref", sqRef);
                selBL.ParentNode.InsertAfter(selBR, selBL);
            }
        }

        private static void ValidateRows(int Row, int Column)
        {
            if (Row < 0 || Row > ExcelPackage.MaxRows - 1)
            {
                throw new ArgumentOutOfRangeException($"Row must not be negative or exceed {ExcelPackage.MaxRows - 1}");
            }

            if (Column < 0 || Column > ExcelPackage.MaxColumns - 1)
            {
                throw new ArgumentOutOfRangeException($"Column must not be negative or exceed {ExcelPackage.MaxColumns - 1}");
            }
        }
        /// <summary>
        /// Split panes at the position in pixels from the top-left corner.
        /// </summary>
        /// <param name="pixelsY">Vertical pixels</param>
        /// <param name="pixelsX">Horizontal pixels</param>
        public void SplitPanesPixels(int pixelsY, int pixelsX)
        {
            if (pixelsY <= 0 && pixelsX <= 0) //Both row and column is zero, remove the panes.
            {
                UnFreezePanes();
                return;
            }
            SetPaneSetting();

            var c = GetTopLeftCell();
            if (pixelsX > 0)
            {
                var styles = _worksheet.Workbook.Styles;
                var normalStyleIx = styles.GetNormalStyleIndex();
                var nf = styles.NamedStyles[normalStyleIx < 0 ? 0 : normalStyleIx].Style.Font;
                var defaultWidth = Convert.ToDouble(FontSize.GetWidthPixels(nf.Name, nf.Size));
                var widthCharRH = c.Row < 1000 ? 3 : c.Row.ToString(CultureInfo.InvariantCulture).Length;
                var margin = 5;
                PaneSettings.XSplit = (Convert.ToDouble(pixelsX) + (defaultWidth * widthCharRH) + margin) * 15D;
            }
            if (pixelsY > 0)
            {
                PaneSettings.YSplit = (pixelsY + _worksheet.DefaultRowHeight / 0.75) * 15D;
            }
            CreateSelectionXml(pixelsY == 0 ? 0 : 1, pixelsX == 0 ? 0 : 1, true);
            Panes = LoadPanes();
            if (pixelsX > 0 && pixelsY > 0)
            {
                var a = new ExcelCellAddress(string.IsNullOrEmpty(TopLeftCell) ? "A1" : TopLeftCell);
                PaneSettings.TopLeftCell = ExcelCellBase.GetAddress(a.Row, a.Column);
            }
        }
        /// <summary>
        /// Split the window at the supplied row/column. 
        /// The split is performed using the current width/height of the visible rows and columns, so any changes to column width or row heights after the split will not effect the split position.
        /// To remove split call this method with zero as value of both paramerters or use <seealso cref="UnFreezePanes"/>
        /// </summary>
        /// <param name="rowsTop">Splits the panes at the coordinate after this visible row. Zero mean no split on row level</param>
        /// <param name="columnsLeft">Splits the panes at the coordinate after this visible column. Zero means no split on column level.</param>
        public void SplitPanes(int rowsTop, int columnsLeft)
        {
            ValidateRows(rowsTop, columnsLeft);
            if (rowsTop == 0 && columnsLeft == 0) //Both row and column is zero, remove the panes.
            {
                UnFreezePanes();
                return;
            }
            SetPaneSetting();

            var c = GetTopLeftCell();
            if (columnsLeft > 0)
            {
                var styles = _worksheet.Workbook.Styles;
                var normalStyleIx = styles.GetNormalStyleIndex();
                var nf = styles.NamedStyles[normalStyleIx < 0 ? 0 : normalStyleIx].Style.Font;
                var defaultWidth = FontSize.GetWidthPixels(nf.Name, nf.Size);
                var widthCharRH = c.Row < 1000 ? 3 : c.Row.ToString(CultureInfo.InvariantCulture).Length;
                var margin = 5;
                PaneSettings.XSplit = (Convert.ToDouble(GetVisibleColumnWidth(c.Column-1, columnsLeft) + (defaultWidth * widthCharRH) + margin)) * 15D;
            }
            if (rowsTop > 0)
            {
                PaneSettings.YSplit = (Convert.ToDouble(GetVisibleRowWidth(c.Row, rowsTop)) + _worksheet.DefaultRowHeight / 0.75) * 15D;
            }
            CreateSelectionXml(rowsTop, columnsLeft, true);
            Panes = LoadPanes();

            var a = new ExcelCellAddress(string.IsNullOrEmpty(TopLeftCell) ? "A1" : TopLeftCell);
            PaneSettings.TopLeftCell = ExcelCellBase.GetAddress(a.Row + rowsTop, a.Column + columnsLeft);
        }

        private void SetPaneSetting()
        {
            if (PaneSettings == null)
            {
                var node = ExcelWorksheetViewPaneSettings.CreatePaneElement(NameSpaceManager, TopNode);
                PaneSettings = new ExcelWorksheetViewPaneSettings(NameSpaceManager, node);
            }
            else
            {
                PaneSettings.State = ePaneState.Split;
            }
        }

        private ExcelCellAddress GetTopLeftCell()
        {
            if (string.IsNullOrEmpty(TopLeftCell))
            {
                if (string.IsNullOrEmpty(PaneSettings?.TopLeftCell))
                {
                    return new ExcelCellAddress();
                }
                else
                {
                    return new ExcelCellAddress(PaneSettings.TopLeftCell);
                }
            }
            else
            {
                return new ExcelCellAddress(TopLeftCell);
            }
        }

        private decimal GetVisibleColumnWidth(int topCol, int cols)
        {
            decimal mdw = _worksheet.Workbook.MaxFontWidth;
            decimal width = 0;
            for (var c = 0; c < cols; c++)
            {
                width += _worksheet.GetColumnWidthPixels(topCol + c, mdw);
            }
            return width;
        }
        private decimal GetVisibleRowWidth(int leftRow, int rows)
        {
            decimal height = 0;
            for (var r = 0; r < rows; r++)
            {
                height += Convert.ToDecimal(_worksheet.GetRowHeight(leftRow + r)) / 0.75M;
            }
            return height;

        }

        private void RemoveSelection()
        {
            //Find selection nodes and remove them            
            XmlNodeList selections = TopNode.SelectNodes(_selectionNodePath, NameSpaceManager);
            foreach (XmlNode sel in selections)
            {
                sel.ParentNode.RemoveChild(sel);
            }
        }
        /// <summary>
        /// Unlock all rows and columns to scroll freely
        /// </summary>
        public void UnFreezePanes()
        {
            string sqRef = SelectedRange, activeCell = ActiveCell;

            XmlElement paneNode = TopNode.SelectSingleNode(_paneNodePath, NameSpaceManager) as XmlElement;
            if (paneNode != null)
            {
                paneNode.ParentNode.RemoveChild(paneNode);
            }
            RemoveSelection();

            PaneSettings = null;
            Panes = new ExcelWorksheetPanes[] { new ExcelWorksheetPanes(NameSpaceManager, TopNode) };

            SelectedRange = sqRef;
            ActiveCell = activeCell;
        }
        #endregion
    }
}
