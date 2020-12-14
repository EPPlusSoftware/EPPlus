/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// Margin setting for a vml drawing
    /// </summary>
    public class ExcelControlMargin
    {
        private ExcelControlWithText _control;
        private XmlHelper _vmlHelper;
        internal ExcelControlMargin(ExcelControlWithText control)
        {
            _control = control;
            _vmlHelper = XmlHelperFactory.Create(control._vmlProp.NameSpaceManager, control._vmlProp.TopNode.ParentNode);
            
            Automatic = _vmlHelper.GetXmlNodeString("@o:insetmode") == "auto";
            var margin = _vmlHelper.GetXmlNodeString("v:textbox/@inset");

            var v = margin.GetCsvPosition(0);
            LeftMargin.SetValue(v);

            v = margin.GetCsvPosition(1);
            TopMargin.SetValue(v);

            v = margin.GetCsvPosition(2);
            RightMargin.SetValue(v);

            v = margin.GetCsvPosition(3);
            BottomMargin.SetValue(v);
        }
        /// <summary>
        /// Sets the margin value and unit of measurement for all margins.
        /// </summary>
        /// <param name="marginValue">Margin value to set for all margins</param>
        /// <param name="unit">The unit to set for all margins. Default <see cref="eMeasurementUnits.Points" /></param>
        public void SetValue(double marginValue, eMeasurementUnits unit=eMeasurementUnits.Points)
        {
            LeftMargin.Value = marginValue;
            TopMargin.Value = marginValue;
            RightMargin.Value = marginValue;
            BottomMargin.Value = marginValue;
            SetUnit(unit);
        }
        /// <summary>
        /// Sets the margin unit of measurement for all margins.
        /// </summary>
        /// <param name="unit">The unit to set for all margins.</param>
        public void SetUnit(eMeasurementUnits unit)
        {
            LeftMargin.Unit = unit;
            TopMargin.Unit = unit;
            RightMargin.Unit = unit;
            BottomMargin.Unit = unit;
        }

        internal void UpdateXml()
        {
            if (Automatic)
            {
                _vmlHelper.SetXmlNodeString("@o:insetmode", "auto");
            }
            else
            {
                _vmlHelper.DeleteNode("@o:insetmode");    //Custom
            }

            if (LeftMargin.Value != 0 && TopMargin.Value != 0 && RightMargin.Value != 0 && BottomMargin.Value != 0)
            {
                var v =
                    LeftMargin.GetValueString() + "," +
                    TopMargin.GetValueString() + "," +
                    RightMargin.GetValueString() + "," +
                    BottomMargin.GetValueString();

                _control.TextBody.LeftInsert = LeftMargin.ToEmu();
                _control.TextBody.TopInsert = TopMargin.ToEmu();
                _control.TextBody.RightInsert = RightMargin.ToEmu(); 
                _control.TextBody.BottomInsert = BottomMargin.ToEmu(); 

                _vmlHelper.SetXmlNodeString("v:textbox/@inset", v);
            }
            else
            {
                _vmlHelper.DeleteNode("v:textbox/@inset");
            }
        }
        /// <summary>
        /// Margin is autiomatic
        /// </summary>
        public bool Automatic
        {
            get;
            set;
        }
        /// <summary>
        /// Left Margin
        /// </summary>
        public ExcelVmlMeasurementUnit LeftMargin
        {
            get;
        } = new ExcelVmlMeasurementUnit();
        /// <summary>
        /// Right Margin
        /// </summary>
        public ExcelVmlMeasurementUnit RightMargin
        {
            get;
        } = new ExcelVmlMeasurementUnit();
        /// <summary>
        /// Top Margin
        /// </summary>
        public ExcelVmlMeasurementUnit TopMargin
        {
            get;
        } = new ExcelVmlMeasurementUnit();
        /// <summary>
        /// Bottom margin
        /// </summary>
        public ExcelVmlMeasurementUnit BottomMargin
        {
            get;
        } = new ExcelVmlMeasurementUnit();
    }
}
