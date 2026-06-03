/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

namespace DrawingRenderer.Constants
{
    internal static class PatternArrays
    {

        internal static readonly short[][] Pct30 = new short[][]
        {
            new short[] { 0, 0, 0, 1 },
            new short[] { 1, 0, 1, 0 },
            new short[] { 0, 1, 0, 0 },
            new short[] { 1, 0, 1, 0 }
        };

        internal static readonly short[][] Pct40 = new short[][]
        {
            new short[] { 0, 0, 0, 1, 0, 1, 0, 1 },
            new short[] { 1, 0, 1, 0, 1, 0, 1, 0 },
            new short[] { 0, 1, 0, 1, 0, 1, 0, 1 },
            new short[] { 1, 0, 1, 0, 1, 0, 1, 0 },
            new short[] { 0, 1, 0, 1, 0, 0, 0, 1 },
            new short[] { 1, 0, 1, 0, 1, 0, 1, 0 },
            new short[] { 0, 1, 0, 1, 0, 1, 0, 1 },
            new short[] { 1, 0, 1, 0, 1, 0, 1, 0 }
        };

        internal static readonly short[][] Pct50 = new short[][]
        {
            new short[] { 1, 1, 1, 0 },
            new short[] { 0, 1, 0, 1 },
            new short[] { 1, 0, 1, 1 },
            new short[] { 0, 1, 0, 1 }
        };

        internal static readonly short[][] Pct60 = new short[][]
        {
            new short[] { 1, 1, 0, 1 },
            new short[] { 0, 1, 1, 1 }
        };

        internal static readonly short[][] Pct70 = new short[][]
        {
            new short[] { 1, 1, 0, 1 },
            new short[] { 1, 1, 1, 1 },
            new short[] { 0, 1, 1, 1 },
            new short[] { 1, 1, 1, 1 }
        };


        internal static readonly short[][] LtHorz = new short[][] {
            new short[] { 1 },
            new short[] { 0 },
            new short[] { 0 },
            new short[] { 0 }
        };

        internal static readonly short[][] LtVert = new short[][] {
            new short[] { 1, 0, 0, 0 }
        };

        internal static readonly short[][] LtUpDiag = new short[][] {
            new short[] { 0, 0, 0, 1 },
            new short[] { 0, 0, 1, 0 },
            new short[] { 0, 1, 0, 0 },
            new short[] { 1, 0, 0, 0 }
        };

        internal static readonly short[][] LtDnDiag = new short[][] {
            new short[] { 1, 0, 0, 0 },
            new short[] { 0, 1, 0, 0 },
            new short[] { 0, 0, 1, 0 },
            new short[] { 0, 0, 0, 1 }
        };

        internal static readonly short[][] DkVert = new short[][] {
            new short[] { 1, 1, 0, 0 }
        };

        internal static readonly short[][] DkHorz = new short[][] {
            new short[] { 1 },
            new short[] { 1 },
            new short[] { 0 },
            new short[] { 0 }
        };


        internal static readonly short[][] DkUpDiag = new short[][] {
            new short[] { 1, 0, 0, 1 },
            new short[] { 0, 0, 1, 1 },
            new short[] { 0, 1, 1, 0 },
            new short[] { 1, 1, 0, 0 }
        };

        internal static readonly short[][] DkDnDiag = new short[][] {
            new short[] { 1, 0, 0, 1 },
            new short[] { 1, 1, 0, 0 },
            new short[] { 0, 1, 1, 0 },
            new short[] { 0, 0, 1, 1 }
        };

        internal static readonly short[][] WdUpDiag = new short[][] {
            new short[] { 1, 1, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 1, 1 },
            new short[] { 0, 0, 0, 0, 0, 1, 1, 0 },
            new short[] { 0, 0, 0, 0, 1, 1, 0, 0 },
            new short[] { 0, 0, 0, 1, 1, 0, 0, 0 },
            new short[] { 0, 0, 1, 1, 0, 0, 0, 0 },
            new short[] { 0, 1, 1, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] WdDnDiag = new short[][] {
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 1, 1, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 1, 1, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 1, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 1, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 1, 1, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 1, 1 }
        };

        internal static readonly short[][] NarVert = new short[][] {
            new short[] { 1, 0 }
        };

        internal static readonly short[][] NarHorz = new short[][] {
            new short[] { 1 },
            new short[] { 0 }
        };


        internal static readonly short[][] Vert = new short[][] {
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] Horz = new short[][] {
            new short[] { 1 },
            new short[] { 0 },
            new short[] { 0 },
            new short[] { 0 },
            new short[] { 0 },
            new short[] { 0 },
            new short[] { 0 },
            new short[] { 0 }
        };

        internal static readonly short[][] DashDnDiag = new short[][] {
            new short[] { 1, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 1, 0, 0, 0, 1, 0, 0 },
            new short[] { 0, 0, 1, 0, 0, 0, 1, 0 },
            new short[] { 0, 0, 0, 1, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] DashUpDiag = new short[][] {
            new short[] { 0, 1, 0, 0, 0, 1, 0, 0 },
            new short[] { 1, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 0, 0, 0, 1 },
            new short[] { 0, 0, 1, 0, 0, 0, 1, 0 }
        };

        internal static readonly short[][] DashHorz = new short[][] {
            new short[] { 0, 0, 0, 0, 1, 1, 1, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 1, 1, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] DashVert = new short[][] {
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 }
        };


        internal static readonly short[][] SmConfetti = new short[][] {
            new short[] { 0, 0, 0, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 1, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 1, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 1, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 1, 0 }
        };

        internal static readonly short[][] LgConfetti = new short[][] {
            new short[] { 0, 0, 0, 0, 0, 0, 1, 1 },
            new short[] { 0, 0, 0, 1, 1, 0, 1, 1 },
            new short[] { 1, 1, 0, 1, 1, 0, 0, 0 },
            new short[] { 1, 1, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 1, 0, 0 },
            new short[] { 1, 0, 0, 0, 1, 1, 0, 1 },
            new short[] { 1, 0, 1, 1, 0, 0, 0, 1 },
            new short[] { 0, 0, 1, 1, 0, 0, 0, 0 }
        };

        internal static readonly short[][] ZigZag = new short[][] {
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 1, 0, 0, 0, 0, 1, 0 },
            new short[] { 0, 0, 1, 0, 0, 1, 0, 0 },
            new short[] { 0, 0, 0, 1, 1, 0, 0, 0 }
        };

        internal static readonly short[][] Wave = new short[][] {
            new short[] { 0, 0, 1, 0, 0, 1, 0, 1 },
            new short[] { 1, 1, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 1, 0, 0, 0 }
        };

        internal static readonly short[][] DiagBrick = new short[][] {
            new short[] { 0, 0, 1, 0, 0, 1, 0, 0 },
            new short[] { 0, 1, 0, 0, 0, 0, 1, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 1, 0 },
            new short[] { 0, 0, 0, 0, 0, 1, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 1, 0, 0, 0 }
        };

        internal static readonly short[][] HorzBrick = new short[][] {
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 1, 1, 1, 1, 1, 1, 1, 1 }
        };

        internal static readonly short[][] Weave = new short[][] {
            new short[] { 1, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 1, 0, 1, 0, 1, 0, 0 },
            new short[] { 0, 0, 1, 0, 0, 0, 1, 0 },
            new short[] { 0, 1, 0, 0, 0, 1, 0, 1 },
            new short[] { 1, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 0, 1, 0, 0 },
            new short[] { 0, 0, 1, 0, 0, 0, 1, 0 },
            new short[] { 0, 1, 0, 1, 0, 0, 0, 1 }
        };

        internal static readonly short[][] Plaid = new short[][] {
            new short[] { 1, 0, 1, 0, 1, 0, 1, 0 },
            new short[] { 0, 1, 0, 1, 0, 1, 0, 1 },
            new short[] { 1, 1, 1, 1, 0, 0, 0, 0 },
            new short[] { 1, 1, 1, 1, 0, 0, 0, 0 },
            new short[] { 1, 1, 1, 1, 0, 0, 0, 0 },
            new short[] { 1, 1, 1, 1, 0, 0, 0, 0 },
            new short[] { 1, 0, 1, 0, 1, 0, 1, 0 },
            new short[] { 0, 1, 0, 1, 0, 1, 0, 1 }
        };


        internal static readonly short[][] Divot = new short[][] {
            new short[] { 0, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] DotGrid = new short[][] {
            new short[] { 1, 0, 1, 0, 1, 0, 1, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] DotDmnd = new short[][] {
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 1, 0, 0, 0, 1, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 1, 0, 0, 0, 1, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] Shingle = new short[][] {
            new short[] { 0, 0, 1, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 1, 1, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 1, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 0, 0, 0, 0, 0, 0, 1, 1 },
            new short[] { 1, 0, 0, 0, 0, 1, 0, 0 },
            new short[] { 0, 1, 0, 0, 1, 0, 0, 0 }
        };

        internal static readonly short[][] Trellis = new short[][] {
            new short[] { 0, 1, 1, 0 },
            new short[] { 1, 1, 1, 1 },
            new short[] { 1, 0, 0, 1 },
            new short[] { 1, 1, 1, 1 }
        };

        internal static readonly short[][] Sphere = new short[][] {
            new short[] { 1, 0, 0, 1, 1, 0, 0, 0 },
            new short[] { 1, 1, 1, 1, 1, 0, 0, 0 },
            new short[] { 1, 1, 1, 1, 1, 0, 0, 0 },
            new short[] { 0, 1, 1, 1, 0, 1, 1, 1 },
            new short[] { 1, 0, 0, 0, 1, 0, 0, 1 },
            new short[] { 1, 0, 0, 0, 1, 1, 1, 1 },
            new short[] { 1, 0, 0, 0, 1, 1, 1, 1 },
            new short[] { 0, 1, 1, 1, 0, 1, 1, 1 }
        };

        internal static readonly short[][] SmGrid = new short[][] {
            new short[] { 1, 1, 1, 1 },
            new short[] { 1, 0, 0, 0 },
            new short[] { 1, 0,  0 },
            new short[] { 1, 0, 0, 0 }
        };

        internal static readonly short[][] LgGrid = new short[][] {
            new short[] { 1, 1, 1, 1, 1, 1, 1, 1 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 0, 0 }
        };

        internal static readonly short[][] SmCheck = new short[][] {
            new short[] { 1, 0, 0, 1 },
            new short[] { 1, 0, 0, 1 },
            new short[] { 0, 1, 1, 0 },
            new short[] { 0, 1, 1, 0 }
        };

        internal static readonly short[][] LgCheck = new short[][] {
            new short[] { 1, 1, 0, 0, 0, 0, 1, 1 },
            new short[] { 1, 1, 0, 0, 0, 0, 1, 1 },
            new short[] { 0, 0, 1, 1, 1, 1, 0, 0 },
            new short[] { 0, 0, 1, 1, 1, 1, 0, 0 },
            new short[] { 0, 0, 1, 1, 1, 1, 0, 0 },
            new short[] { 0, 0, 1, 1, 1, 1, 0, 0 },
            new short[] { 1, 1, 0, 0, 0, 0, 1, 1 },
            new short[] { 1, 1, 0, 0, 0, 0, 1, 1 }
        };

        internal static readonly short[][] OpenDmnd = new short[][] {
            new short[] { 0, 1, 0, 0, 0, 1, 0, 0 },
            new short[] { 1, 0, 0, 0, 0, 0, 1, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 1 },
            new short[] { 1, 0, 0, 0, 0, 0, 1, 0 },
            new short[] { 0, 1, 0, 0, 0, 1, 0, 0 },
            new short[] { 0, 0, 1, 0, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 1, 0, 1, 0, 0, 0 }
        };

        internal static readonly short[][] SolidDmnd = new short[][] {
            new short[] { 0, 1, 1, 1, 1, 1, 0, 0 },
            new short[] { 0, 0, 1, 1, 1, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 0, 0, 0, 0, 0 },
            new short[] { 0, 0, 0, 1, 0, 0, 0, 0 },
            new short[] { 0, 0, 1, 1, 1, 0, 0, 0 },
            new short[] { 0, 1, 1, 1, 1, 1, 0, 0 },
            new short[] { 1, 1, 1, 1, 1, 1, 1, 0 }
        };

    }
}
