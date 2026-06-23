using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Helpers
{
    /// <summary>
    /// Horizontal text alignment
    /// </summary>
    public enum ExcelHorizontalAlignment
    {
        /// <summary>
        /// General aligned
        /// </summary>
        General,
        /// <summary>
        /// Left aligned
        /// </summary>
        Left,
        /// <summary>
        /// Center aligned
        /// </summary>
        Center,
        /// <summary>
        /// The horizontal alignment is centered across multiple cells
        /// </summary>
        CenterContinuous,
        /// <summary>
        /// Right aligned
        /// </summary>
        Right,
        /// <summary>
        /// The value of the cell should be filled across the entire width of the cell.
        /// </summary>
        Fill,
        /// <summary>
        /// Each word in each line of text inside the cell is evenly distributed across the width of the cell
        /// </summary>
        Distributed,
        /// <summary>
        /// The horizontal alignment is justified to the Left and Right for each row.
        /// </summary>
        Justify
    }
    /// <summary>
    /// Vertical text alignment
    /// </summary>
    public enum ExcelVerticalAlignment
    {
        /// <summary>
        /// Top aligned
        /// </summary>
        Top,
        /// <summary>
        /// Center aligned
        /// </summary>
        Center,
        /// <summary>
        /// Bottom aligned
        /// </summary>
        Bottom,
        /// <summary>
        /// Distributed. Each line of text inside the cell is evenly distributed across the height of the cell
        /// </summary>
        Distributed,
        /// <summary>
        /// Justify. Each line of text inside the cell is evenly distributed across the height of the cell
        /// </summary>
        Justify
    }
    /// <summary>
    /// The reading order
    /// </summary>
    public enum ExcelReadingOrder
    {
        /// <summary>
        /// Reading order is determined by the first non-whitespace character
        /// </summary>
        ContextDependent = 0,
        /// <summary>
        /// Left to Right
        /// </summary>
        LeftToRight = 1,
        /// <summary>
        /// Right to Left
        /// </summary>
        RightToLeft = 2
    }
    /// <summary>
    /// Fill pattern
    /// </summary>
    public enum ExcelFillStyle
    {
        /// <summary>
        /// No fill
        /// </summary>
        None,
        /// <summary>
        /// A solid fill
        /// </summary>
        Solid,
        /// <summary>
        /// Dark gray  <para/>
        /// Excel name: 75% Gray
        /// </summary>
        DarkGray,
        /// <summary>
        /// Medium gray <para/>
        /// Excel name: 50% Gray
        /// </summary>
        MediumGray,
        /// <summary>
        /// Light gray <para/>
        /// Excel name: 25% Gray
        /// </summary>
        LightGray,
        /// <summary>
        /// Grayscale of 0.125, 1/8 <para/>
        /// Excel name: 12.5% Gray
        /// </summary>
        Gray125,
        /// <summary>
        /// Grayscale of 0.0625, 1/16 <para/>
        /// Excel name: 6.25% Gray
        /// </summary>
        Gray0625,
        /// <summary>
        /// Dark vertical <para/>
        /// Excel name: Vertical Stripe
        /// </summary>
        DarkVertical,
        /// <summary>
        /// Dark horizontal <para/>
        /// Excel name: Horizontal Stripe
        /// </summary>
        DarkHorizontal,
        /// <summary>
        /// Dark down <para/>
        /// Excel name: Reverse Diagonal Stripe
        /// </summary>
        DarkDown,
        /// <summary>
        /// Dark up <para/>
        /// Excel name: Diagonal Stripe
        /// </summary>
        DarkUp,
        /// <summary>
        /// Dark grid <para/>
        /// Excel name: Diagonal Crosshatch
        /// </summary>
        DarkGrid,
        /// <summary>
        /// Dark trellis <para/>
        /// Excel name: Thick Diagonal Crosshatch
        /// </summary>
        DarkTrellis,
        /// <summary>
        /// Light vertical <para/>
        /// Excel name: Thin Vertical Stripe
        /// </summary>
        LightVertical,
        /// <summary>
        /// Light horizontal <para/>
        /// Excel name: Thin Horizontal Stripe
        /// </summary>
        LightHorizontal,
        /// <summary>
        /// Light down <para/>
        /// Excel name: Thin Reverse Diagonal Stripe
        /// </summary>
        LightDown,
        /// <summary>
        /// Light up <para/>
        /// Excel name: Thin Diagonal Stripe
        /// </summary>
        LightUp,
        /// <summary>
        /// Light grid <para/>
        /// Excel name: Thin Horizontal Crosshatch
        /// </summary>
        LightGrid,
        /// <summary>
        /// Light trellis <para/>
        /// Excel name: Thin Diagonal Crosshatch
        /// </summary>
        LightTrellis
    }
    /// <summary>
    /// BulletType of gradient fill
    /// </summary>
    public enum ExcelFillGradientType
    {
        /// <summary>
        /// No gradient fill. 
        /// </summary>
        None,
        /// <summary>
        /// Linear gradient type. Linear gradient type means that the transition from one color to the next is along a line.
        /// </summary>
        Linear,
        /// <summary>
        /// Path gradient type. Path gradient type means the that the transition from one color to the next is a rectangle, defined by coordinates.
        /// </summary>
        Path
    }
    /// <summary>
    /// Border line style
    /// </summary>
    public enum ExcelBorderStyle
    {
        /// <summary>
        /// No border style
        /// </summary>
        None,
        /// <summary>
        /// Hairline
        /// </summary>
        Hair,
        /// <summary>
        /// Dotted
        /// </summary>
        Dotted,
        /// <summary>
        /// Dash Dot
        /// </summary>
        DashDot,
        /// <summary>
        /// Thin single line
        /// </summary>
        Thin,
        /// <summary>
        /// Dash Dot Dot
        /// </summary>
        DashDotDot,
        /// <summary>
        /// Dashed
        /// </summary>
        Dashed,
        /// <summary>
        /// Dash Dot Dot, medium thickness
        /// </summary>
        MediumDashDotDot,
        /// <summary>
        /// Dashed, medium thickness
        /// </summary>
        MediumDashed,
        /// <summary>
        /// Dash Dot, medium thickness
        /// </summary>
        MediumDashDot,
        /// <summary>
        /// Single line, Thick
        /// </summary>
        Thick,
        /// <summary>
        /// Single line, medium thickness
        /// </summary>
        Medium,
        /// <summary>
        /// Double line
        /// </summary>
        Double,
        /// <summary>
        /// Slanted Dash Dot
        /// </summary>
        SlantDashDot
    }
}
