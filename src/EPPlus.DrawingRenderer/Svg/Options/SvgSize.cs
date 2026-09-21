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

using EPPlus.Fonts.OpenType.Utils;
using System.Globalization;

namespace EPPlus.DrawingRenderer.Svg
{
    /// <summary>
    /// Specifies the CSS/SVG length unit used when rendering an <see cref="SvgDimension"/>
    /// as an SVG presentation attribute value (e.g. the <c>width</c>/<c>height</c> attributes
    /// on the root <c>&lt;svg&gt;</c> element).
    /// </summary>
    public enum eSvgUnit
    {
        /// <summary>
        /// Use the drawing dimensions, this is the default for <see cref="SvgSize.Width"/> and <see cref="SvgSize.Height"/>.
        /// </summary>
        UseDrawingDimension,
        /// <summary>
        /// No unit suffix is written; the value is emitted as a bare number (e.g. <c>"600"</c>).
        /// Per the SVG specification this is equivalent to <see cref="Px"/> for the outermost
        /// <c>&lt;svg&gt;</c> element, but the two are written as different strings. This is the
        /// default used by the implicit <c>double</c> conversion, preserving EPPlus's existing
        /// (pre-unit) SVG export output.
        /// </summary>
        None,

        /// <summary>CSS pixels, written with an explicit <c>px</c> suffix (e.g. <c>"600px"</c>).</summary>
        Px,

        /// <summary>Points (1pt = 1/72in), written with a <c>pt</c> suffix.</summary>
        Pt,

        /// <summary>Picas (1pc = 12pt), written with a <c>pc</c> suffix.</summary>
        Pc,

        /// <summary>Inches, written with an <c>in</c> suffix.</summary>
        In,

        /// <summary>Centimeters, written with a <c>cm</c> suffix.</summary>
        Cm,

        /// <summary>Millimeters, written with a <c>mm</c> suffix.</summary>
        Mm,

        /// <summary>
        /// Font-relative unit (the element's font size), written with an <c>em</c> suffix.
        /// Meaningful mainly when the SVG is embedded inline in HTML and inherits a font context.
        /// </summary>
        Em,

        /// <summary>
        /// Font-relative unit (the font's x-height), written with an <c>ex</c> suffix.
        /// Meaningful mainly when the SVG is embedded inline in HTML and inherits a font context.
        /// </summary>
        Ex,

        /// <summary>
        /// A percentage of the containing viewport, written with a <c>%</c> suffix (e.g. <c>"100%"</c>).
        /// Requires a <c>viewBox</c> on the root element to preserve aspect ratio as the container resizes.
        /// </summary>
        Percent,
        /// <summary>
        /// A percentage of the viewport width, written with a <c>vw</c> suffix (e.g. <c>"100vw"</c>).
        /// </summary>
        Vw,
        /// <summary>
        /// A percentage of the viewport height, written with a <c>vh</c> suffix (e.g. <c>"100vh"</c>).
        /// </summary>
        Vh,
        /// <summary>
        /// A percentage of the viewport's smaller dimension, written with a <c>vmin</c> suffix (e.g. <c>"100vmin"</c>).
        /// </summary>
        Vmin,
        /// <summary>
        /// A percentage of the viewport's larger dimension, written with a <c>vmax</c> suffix (e.g. <c>"100vmax"</c>).
        /// </summary>
        Vmax,
        /// <summary>
        /// The attribute will be removed from the output. This specifieds that the width or height attribute will be removed from the svg output.
        /// See <see cref="SvgDimension.Removed()"/> and <see cref="SvgDimension.IsRemoved"/>
        /// </summary>
        Removed
    }

    /// <summary>
    /// Represents a single dimension (width or height) for SVG export, as either a numeric value
    /// with an explicit or implicit unit (<see cref="eSvgUnit"/>) or the SVG2 <c>auto</c> keyword.
    /// </summary>
    /// <remarks>
    /// Instances are created via the static factory methods (<see cref="Default"/>, <see cref="Px"/>,
    /// <see cref="Pt"/>, etc.) or the static <see cref="Auto"/> field, and are rendered to an SVG
    /// attribute string with <see cref="ToAttributeString"/>. A plain <see cref="double"/> converts
    /// implicitly to a unitless (<see cref="eSvgUnit.None"/>) value, matching EPPlus's historical SVG
    /// export output, so existing call sites that pass raw pixel numbers are unaffected.
    /// </remarks>
    public class SvgDimension
    {
        string _attributeName;
        internal SvgDimension()
        {
            SetUseDrawingDimension();
        }
        /// <summary>The unit this dimension is expressed in.</summary>
        public eSvgUnit Unit { get; private set; }

        /// <summary>
        /// The numeric magnitude of this dimension. Unused (always <c>0</c>) when
        /// <see cref="Unit"/> is <see cref="eSvgUnit.Auto"/>.
        /// </summary>
        public double? Value { get; private set; }
        /// <summary>
        /// It true, the attribute will be removed from the svg output. See <see cref="SvgDimension.Remove()"/> />
        /// </summary>
        public bool IsRemoved { get => Unit == eSvgUnit.Removed;  }
        /// <summary>
        /// Creates a unitless dimension (e.g. <c>"600"</c>) — the SVG default, equivalent in every
        /// conformant renderer to <see cref="Px"/> but written without a unit suffix. This matches
        /// EPPlus's pre-existing SVG export output and is what the implicit <see cref="double"/>
        /// conversion produces.
        /// </summary>
        /// <param name="value">The value in CSS pixels.</param>
        public void SetDefault(double value) => SetUnit(eSvgUnit.None, value);
        /// <summary>
        /// The value from the drawing will be used. This is the default for <see cref="SvgSize.Width"/> and <see cref="SvgSize.Height"/>.
        /// </summary>
        public void SetUseDrawingDimension() => SetUnit(eSvgUnit.UseDrawingDimension);

        /// <summary>Creates a dimension in CSS pixels, written with an explicit <c>px</c> suffix.</summary>
        /// <param name="value">The value in CSS pixels.</param>
        public void SetPixels(double value) => SetUnit(eSvgUnit.Px, value);

        /// <summary>Creates a dimension in points (1pt = 1/72in).</summary>
        /// <param name="value">The value in points.</param>
        public void SetPoints(double value) => SetUnit(eSvgUnit.Pt, value);

        /// <summary>Creates a dimension in picas (1pc = 12pt).</summary>
        /// <param name="value">The value in picas.</param>
        public void SetPicas(double value) => SetUnit(eSvgUnit.Pc, value);

        /// <summary>Creates a dimension in inches.</summary>
        /// <param name="value">The value in inches.</param>
        public void SetInches(double value) => SetUnit(eSvgUnit.In, value);

        /// <summary>Creates a dimension in centimeters.</summary>
        /// <param name="value">The value in centimeters.</param>
        public void SetCentimeters(double value) => SetUnit(eSvgUnit.Cm, value);

        /// <summary>Creates a dimension in millimeters.</summary>
        /// <param name="value">The value in millimeters.</param>
        public void SetMillimeters(double value) => SetUnit(eSvgUnit.Mm, value);

        /// <summary>Creates a font-relative dimension in <c>em</c> units.</summary>
        /// <param name="value">The value in em units.</param>
        public void SetEms(double value) => SetUnit(eSvgUnit.Em, value);

        /// <summary>Creates a font-relative dimension in <c>ex</c> units.</summary>
        /// <param name="value">The value in ex units.</param>
        public void SetEx(double value) => SetUnit(eSvgUnit.Ex, value);
        /// <summary>
        /// Sets this dimension in viewport-width units (1vw = 1% of the browser viewport width).
        /// Resolves against the browser window, not the element's containing block — prefer
        /// <see cref="SetPercent"/> for "fill my container" sizing; use this only when the chart
        /// is meant to fill the actual browser viewport.
        /// </summary>
        /// <param name="value">The value in vw units.</param>
        public void SetViewportWidth(double value) => SetUnit(eSvgUnit.Vw, value);

        /// <summary>Sets this dimension in viewport-height units. See <see cref="SetViewportWidth"/> remarks.</summary>
        /// <param name="value">The value in vh units.</param>
        public void SetViewportHeight(double value) => SetUnit(eSvgUnit.Vh, value);

        /// <summary>Sets this dimension in viewport-min units (1% of the smaller viewport dimension).
        /// See <see cref="SetViewportWidth"/> remarks.</summary>
        /// <param name="value">The value in vmin units.</param>
        public void SetViewportMin(double value) => SetUnit(eSvgUnit.Vmin, value);

        /// <summary>Sets this dimension in viewport-max units (1% of the larger viewport dimension).
        /// See <see cref="SetViewportWidth"/> remarks.</summary>
        /// <param name="value">The value in vmax units.</param>
        public void SetViewportMax(double value) => SetUnit(eSvgUnit.Vmax, value);
        /// <summary>
        /// Creates a percentage dimension, resolved by the renderer against the containing viewport.
        /// The root element's <c>viewBox</c> should be set so the aspect ratio is preserved as the
        /// container resizes.
        /// </summary>
        /// <param name="value">The percentage value, e.g. <c>100</c> for <c>"100%"</c>.</param>
        public void SetPercent(double value) => SetUnit(eSvgUnit.Percent, value);
        private void SetUnit(eSvgUnit unit, double value=double.NaN)
        {
            Unit = unit;
            Value = value;
        }
        public void Remove()
        {
            Unit = eSvgUnit.Removed;
            Value = null;
        }
        private static readonly Dictionary<eSvgUnit, string> _suffix = new()
        {
            [eSvgUnit.Px] = "px",
            [eSvgUnit.Pt] = "pt",
            [eSvgUnit.Pc] = "pc",
            [eSvgUnit.In] = "in",
            [eSvgUnit.Cm] = "cm",
            [eSvgUnit.Mm] = "mm",
            [eSvgUnit.Em] = "em",
            [eSvgUnit.Ex] = "ex",
            [eSvgUnit.Percent] = "%",
            [eSvgUnit.Vw] = "vw",
            [eSvgUnit.Vh] = "vh",
            [eSvgUnit.Vmin] = "vmin",
            [eSvgUnit.Vmax] = "vmax"
        };
        /// <summary>
        /// Operator to implicitly convert a <c>double</c> to an <see cref="SvgDimension"/> with no unit suffix (pixels)
        /// </summary>
        /// <param name="px"></param>
        public static implicit operator SvgDimension(double px)
        {
            var d = new SvgDimension();
            d.SetDefault(px);
            return d;
        }
        /// <summary>
        /// Renders this dimension as the string to write into an SVG length-valued attribute
        /// (e.g. <c>width</c>/<c>height</c> on the root <c>&lt;svg&gt;</c> element), such as
        /// <c>"600"</c>, <c>"600px"</c>, <c>"210mm"</c>, <c>"100%"</c>, or <c>"auto"</c>.
        /// </summary>
        internal string ToAttributeString(string attributeName, double drawingDimensionPx)
        {
            if (IsRemoved) return "";
            return Unit switch
            {

                eSvgUnit.UseDrawingDimension => attributeName + "=\"" + drawingDimensionPx.ToString(CultureInfo.InvariantCulture) + "\"",
                _ => attributeName + "=\"" + Value.GetValueOrDefault().ToString(CultureInfo.InvariantCulture) + (_suffix.TryGetValue(Unit, out var s) ? s : "") + "\""
            };
        }
    }
    public class SvgSize
    {
        /// <summary>
        /// Overrides the width for ouput the svg image. If not specified the drawings dimension will be used.
        /// </summary>
        public SvgDimension Width { get; set; } = new SvgDimension();
        /// <summary>
        /// Overrides the width for ouput the svg image. If not specified the drawings dimension will be used.
        /// </summary>
        public SvgDimension Height { get; set; } = new SvgDimension();
    }
}

