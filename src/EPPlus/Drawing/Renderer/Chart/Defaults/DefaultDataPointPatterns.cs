using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Renderer.Chart.Defaults
{
    internal static class DefaultDataPointPatterns
    {
        /// <summary>
        /// Excel built-in fallback colors when no style and no theme
        /// </summary>
        internal readonly static Color[] FallBackColors =
        { Color.FromArgb(255, 79, 129, 189),
          Color.FromArgb(255, 192, 80, 77),
          Color.FromArgb(255, 155, 187, 89),
          Color.FromArgb(255, 128, 100, 162),
          Color.FromArgb(255, 75, 172, 198),
          Color.FromArgb(255, 247, 150, 70)
        };

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataPointCount"></param>
        /// <param name="origColors">usually the 6 accent colors</param>
        /// <returns></returns>
        internal static List<List<Color>> GetLegacyColorVariations(int dataPointCount, List<Color> origColors)
        {

            var numVariations = dataPointCount / origColors.Count();

            //KEY TAKE AWAYS TO MAKE FUNC WORK
            //VariationIndex begins at 1 NOT 0
            //highestVariation INCREASED by 1 when REACHING 6 not the start of the next "accent1" 
            if (dataPointCount % origColors.Count() == 0)
            {
                numVariations += 1;
            }

            List<List<Color>> outputColors = new List<List<Color>>();
            for (int i = 1; i < numVariations; i++)
            {
                List<Color> currentVariations = new List<Color>();
                for (int j = 0; j < origColors.Count; j++)
                {
                    var color = origColors[j];
                    var tintValue = (-70d + 140d * (i / (numVariations + 1d))) / 100d;

                    var variation = OfficeOpenXml.Utils.TypeConversion.ColorConverter.ApplyTintDrawing(color, tintValue);

                    currentVariations.Add(variation);
                }
                outputColors.Add(currentVariations);
            }

            return outputColors;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="variationIndex">0-based variation index</param>
        /// <param name="numOfVariations">Total number of variations (Should be +1 if datapoints evenly dividable with num original colors)</param>
        /// <param name="origColor"></param>
        /// <returns></returns>
        internal static Color GetSingleLegacyColorVariation(int variationIndex, int numOfVariations, Color origColor)
        {
            //legacy uses 1-based variation index
            variationIndex += 1;

            var tintValue = (-70d + 140d * (variationIndex / (numOfVariations + 1d))) / 100d;

            return OfficeOpenXml.Utils.TypeConversion.ColorConverter.ApplyTintDrawing(origColor, tintValue);
        }
    }
}
