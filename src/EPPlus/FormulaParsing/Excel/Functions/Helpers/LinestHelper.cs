/*************************************************************************************************
 Required Notice: Copyright (C) EPPlus Software AB. 
 This software is licensed under PolyForm Noncommercial License 1.0.0 
 and may only be used for noncommercial purposes 
 https://polyformproject.org/licenses/noncommercial/1.0.0/

 A commercial license to use this software can be purchased at https://epplussoftware.com
*************************************************************************************************
 Date               Author                       Change
*************************************************************************************************
 05/07/2023         EPPlus Software AB         Implemented function
*************************************************************************************************/
using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Text;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
//using System.ComponentModel.DataAnnotations;
namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class LinestHelper
    {
        public static InMemoryRange MultipleRegResult(double[] knownYs, double[][] xRangeList, bool constVar, bool stats, bool logest)
        {

            //Adding a column of ones to account for the intercept.
            //This column will be the first column in the matrix, so all columns are shifted to the right to make space.
            if (constVar)
            {
                for (var r = 0; r < xRangeList.Count(); r++)
                {
                    for (var c = xRangeList[r].Count() - 1; c > 0; c--)
                    {
                        xRangeList[r][c] = xRangeList[r][c - 1];
                    }
                }
                for (var r = 0; r < xRangeList.Count(); r++)
                {
                    xRangeList[r][0] = 1d;
                }
            }

            var width = xRangeList[0].Count();
            var height = xRangeList.Count();

            //Gaussian elimination to get rank and which columns to be dropped due to collinearity.
            //Choosing between two collinear columns, the later is chosen.
            var dropCols = MatrixHelper.GaussRank(xRangeList, constVar);
            var xRangeListCopy = xRangeList;

            //columns are dropped based of off dropCols. The drop columns are represented as zero in the coefficients and standard error results.
            if (dropCols.Count() > 0) xRangeList = MatrixHelper.RemoveColumns(xRangeList, dropCols);

            //LOGEST can be expressed as EXP(LINEST(LN(yRange), xRange))
            //We simply take the natural logarithm of the predictor value and apply the function above to retrieve the coefficients for a non-linear regression
            if (logest)
            {
                for (var i = 0; i < knownYs.Length; i++)
                {
                    knownYs[i] = Math.Log(knownYs[i]);
                }
            }

            //GetSlope calculates the OLS estimator
            var multipleRegressionSlopes = GetSlope(xRangeList, knownYs, constVar, stats, logest, out bool matrixIsSingular);

            //If LOGEST, we transform the linear results according to the formula LOGEST = EXP(LINEST(LN(y), x))
            if (logest)
            {
                for (var i = 0; i < multipleRegressionSlopes.Length; i++)
                {
                    multipleRegressionSlopes[i] = Math.Exp(multipleRegressionSlopes[i]);
                }
            }

            var betaCoefficients = ReverseCoefficientOrder(multipleRegressionSlopes, dropCols, constVar, logest);

            if (!stats)
            {
                //If no stats, we just print out the coefficients and return the range
                var resultRange = new InMemoryRange(1, (short)(betaCoefficients.Count()));
                for (var i = 0; i < betaCoefficients.Count(); i++)
                {
                    resultRange.SetValue(0, i, betaCoefficients[i]);
                }
                return resultRange;
            }
            else
            {
                //Print coefficients
                var resultRangeStats = new InMemoryRange(5, (short)betaCoefficients.Count());
                for (var i = 0; i < betaCoefficients.Count(); i++)
                {
                    resultRangeStats.SetValue(0, i, betaCoefficients[i]);
                }

                //Get the sum statistics: Regression sum of squares and residual sum of squares
                GetSums(betaCoefficients, xRangeListCopy, knownYs, logest, constVar, out double ssreg, out double ssresid);
                
                //Get the remaining statistics
                var rSquared = ssreg / (ssreg + ssresid);
                var df = Math.Max(height - width + dropCols.Count(), 0d); //Adjust df when columns are dropped --> For every dropped column, +1 to degrees of freedom
                var standardErrorEstimate = (df != 0d) ? Math.Sqrt(ssresid / df) : 0d;
                object fStatistic = 0d;
                if (df != 0)
                {
                    fStatistic = (constVar) ? (ssreg / (width - 1 - dropCols.Count())) / (ssresid / df) : (ssreg / width) / (ssresid / df);
                }
                else
                {
                    fStatistic = ExcelErrorValue.Create(eErrorType.Num);
                }

                //Get the standard errors for the coefficients
                var standardErrorArray = GetStandardErrors(df, ssresid, height, width, dropCols, xRangeList); 
                //Print standard errors on the second row in the result range
                if (constVar)
                {
                    resultRangeStats.SetValue(1, xRangeListCopy[0].Count() - 1, standardErrorArray[standardErrorArray.Count() - 1]);
                }
                else
                {
                    resultRangeStats.SetValue(1, xRangeListCopy[0].Count(), ExcelErrorValue.Create(eErrorType.NA));
                }

                int pos3 = 0;
                for (var i = standardErrorArray.Count() - 1; i >= 0; i--)
                {
                    resultRangeStats.SetValue(1, pos3++, standardErrorArray[i]);
                }

                //Print #N/A in the result range
                if (constVar)
                {
                    for (var col = 2; col < width; col++)
                    {
                        for (var row = 2; row < 5; row++)
                        {
                            resultRangeStats.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                        }
                    }
                }
                else
                {
                    for (var col = 2; col < width + 1; col++)
                    {
                        for (var row = 2; row < 5; row++)
                        {
                            resultRangeStats.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                        }
                    }
                }

                //Print out the remaining statistics
                resultRangeStats.SetValue(2, 0, rSquared);
                resultRangeStats.SetValue(2, 1, standardErrorEstimate);
                resultRangeStats.SetValue(3, 0, fStatistic);
                resultRangeStats.SetValue(3, 1, df);
                resultRangeStats.SetValue(4, 0, ssreg);
                resultRangeStats.SetValue(4, 1, ssresid);
                return resultRangeStats;
            }
        }

        private static double[] GetSlope(double[][] xValues, double[] yValues, bool constVar, bool stats, bool logest, out bool matrixIsSingular)
        {
            var width = xValues[0].Count();
            var height = xValues.Count();
            var xT = MatrixHelper.TransposeMatrix(xValues, height, width);
            var xTdotX = MatrixHelper.Multiply(xT, xValues);
            var myInverse = MatrixHelper.Inverse(xTdotX);
            var dotProduct = MatrixHelper.Multiply(myInverse, xT);
            double[][] yValuesJagged = yValues.Select(yVal => new double[] { yVal }).ToArray();

            //b = (X'X)^-1 * X' * Y
            var b = MatrixHelper.Multiply(dotProduct, yValuesJagged);
            double[] bArray = b.Select(innerList => innerList[0]).ToArray();
            matrixIsSingular = (MatrixHelper.GetDeterminant(xTdotX) < 1E-8) ? true : false; //This threshold could be investigated further

            if (!constVar)
            {
                double[] extendedB = new double[bArray.Count() + 1];
                for (var i = 0; i < bArray.Count(); i++)
                {
                    extendedB[i + 1] = bArray[i];
                }
                //Since we log the linear intercept, we can put the LOGEST intercept since ln(0) = 1
                extendedB[0] = 0d;
                return extendedB;
            }
            return bArray;
        }

        public static InMemoryRange LinearRegResult(double[] knownXs, double[] knownYs, bool constVar, bool stats, bool logest)
        {
            var knownYsCopy = knownYs.ToList();
            if (logest)
            {
                for (var i = 0; i < knownYs.Count(); i++)
                {
                    knownYs[i] = Math.Log(knownYs[i]);
                }
            }

            var averageY = knownYs.Average();
            var averageX = knownXs.Average();

            double nominator = 0d;
            double denominator = 0d;
            double xDiff = 0d;
            double yDiff = 0d;
            double estimatedDiff = 0d;
            double ssr = 0d;
            double sst = 0d;
            var df = 0d;
            var v1 = 0d;
            var v2 = 0d;
            var fStatistics = 0d;

            for (var i = 0; i < knownYs.Count(); i++)
            {
                var y = knownYs[i];
                var x = knownXs[i];

                if (constVar)
                {
                    nominator += (x - averageX) * (y - averageY);
                    denominator += (x - averageX) * (x - averageX);
                }
                else
                {
                    nominator += x * y;
                    denominator += Math.Pow(x, 2);
                }

            }

            var m = (denominator != 0) ? nominator / denominator : 0d;
            var b = (constVar) ? averageY - (m * averageX) : 0d;

            //LOGEST can be expressed as EXP(LINEST(LN(y), x))
            if (logest) m = Math.Exp(m);
            if (logest) b = Math.Exp(b);

            if (stats)
            {
                for (var i = 0; i < knownXs.Count(); i++)
                {
                    var x = knownXs[i];
                    var y = knownYs[i];

                    //LOGEST uses the same statistics as LINEST, but with logged y-values. We remove the EXP to get correct statistics (correct according to excel)
                    var estimatedY = (logest) ? Math.Log(m) * x + Math.Log(b) : m * x + b; //LINEST formula

                    if (constVar)
                    {
                        estimatedDiff += Math.Pow(y - estimatedY, 2);
                        xDiff += Math.Pow(x - averageX, 2);
                        yDiff += Math.Pow(y - estimatedY, 2);
                        ssr += Math.Pow(estimatedY - averageY, 2);
                        sst += Math.Pow(y - averageY, 2);
                    }
                    else
                    {
                        estimatedDiff += Math.Pow(y - estimatedY, 2);
                        xDiff += Math.Pow(x, 2);
                        yDiff = Math.Pow(y - estimatedY, 2);
                        ssr += Math.Pow(estimatedY, 2);
                        sst += Math.Pow(y, 2);
                    }

                }

                var errorVariance = yDiff / (knownXs.Count() - 2);
                if (!constVar) errorVariance = yDiff / (knownXs.Count() - 1);

                var standardErrorM = (constVar) ? Math.Sqrt(1d / (knownXs.Count() - 2d) * estimatedDiff / xDiff) :
                                                  Math.Sqrt(1d / (knownXs.Count() - 1d) * estimatedDiff / xDiff);

                object standardErrorB = Math.Sqrt(errorVariance) * Math.Sqrt(1d / knownXs.Count() + Math.Pow(averageX, 2) / xDiff);
                if (!constVar) standardErrorB = ExcelErrorValue.Create(eErrorType.NA);

                var rSquared = ssr / sst;
                var standardErrorEstimateY = (!constVar) ? SEHelper.GetStandardError(knownXs, knownYs, true) :
                                                          SEHelper.GetStandardError(knownXs, knownYs, false);
                var ssreg = ssr;
                var ssresid = (constVar) ? yDiff : (sst - ssr);

                if (constVar)
                {
                    df = knownXs.Count() - 2;
                    v1 = knownXs.Count() - df - 1;
                    v2 = df;
                    fStatistics = (ssr / v1) / (yDiff / v2);
                }
                else
                {
                    df = knownXs.Count() - 1;
                    v1 = knownXs.Count() - df;
                    v2 = df;
                    fStatistics = ssr / (ssresid / (knownXs.Count() - 1));
                }

                var resultRangeStats = new InMemoryRange(5, 2);
                resultRangeStats.SetValue(0, 0, m);
                resultRangeStats.SetValue(0, 1, b);
                resultRangeStats.SetValue(1, 0, standardErrorM);
                resultRangeStats.SetValue(1, 1, standardErrorB);
                resultRangeStats.SetValue(2, 0, rSquared);
                resultRangeStats.SetValue(2, 1, standardErrorEstimateY);
                resultRangeStats.SetValue(3, 0, fStatistics);
                resultRangeStats.SetValue(3, 1, df);
                resultRangeStats.SetValue(4, 0, ssreg);
                resultRangeStats.SetValue(4, 1, ssresid);
                return resultRangeStats;
            }

            var resultRangeNormal = new InMemoryRange(1, 2);
            resultRangeNormal.SetValue(0, 0, m);
            resultRangeNormal.SetValue(0, 1, b);
            return resultRangeNormal;

        }
        internal static InMemoryRange ExecuteLinest(IRangeInfo rangeX, IRangeInfo rangeY, bool constVar, bool stats, bool logest, out eErrorType? error)
        {
            bool multipleXranges = false;
            bool columnArray = false;
            bool rowArray = false;
            var xColumns = rangeX.Size.NumberOfCols;
            var yColumns = rangeY.Size.NumberOfCols;
            var xRows = rangeX.Size.NumberOfRows;
            var yRows = rangeY.Size.NumberOfRows;
            error = null;

            if ((xRows != yRows && xColumns == yColumns)
                || (xColumns != yColumns && xRows == yRows))
            {
                multipleXranges = true;
            }
            else
            {
                if (xRows != yRows || xColumns != yColumns)
                {
                    error = eErrorType.Ref;
                    return null;
                };
            }

            RangeFlattener.GetNumericPairLists(rangeX, rangeY, !multipleXranges, out List<double> knownXsList, out List<double> knownYsList);
            //y values cant be zero or negative since we have to take the logarithm of the y-values to find a solution.
            if (logest)
            {
                for (var i = 0; i < knownYsList.Count(); i++)
                {
                    if (knownYsList[i] <= 0)
                    {
                        error = eErrorType.Num;
                        return null;
                    };
                }
            }

            var knownXs = MatrixHelper.ListToArray(knownXsList);
            var knownYs = MatrixHelper.ListToArray(knownYsList);
            var r = 0;
            var c = 0;

            if (multipleXranges)
            {
                if (multipleXranges && xColumns != yColumns)
                {
                    columnArray = true;

                    r = xRows;
                    c = xColumns;
                }
                else if (multipleXranges && xRows != yRows)
                {
                    rowArray = true;
                    r = xColumns;
                    c = xRows;
                }
            }
            else
            {
                r = knownXs.Count();
                c = 1;
            }
            if (multipleXranges && constVar)
            {
                c += 1; //This is because we need to add a vector of ones to the matrix in order to account for the intercept
            }

            double[][] xRanges = MatrixHelper.CreateMatrix(r, c);

            if (columnArray)
            {
                var counter = 0;
                var delimiter = (constVar) ? xRanges[0].Count() - 1 : xRanges[0].Count();
                for (var i = 0; i < xRanges.Count(); i++)
                {
                    for (var j = 0; j < delimiter; j++)
                    {
                        xRanges[i][j] = knownXs[counter];
                        counter += 1;
                    }
                }
            }

            else if (rowArray)
            {
                //This shifts data thats row-based to column-based.
                var counter = 0;
                var delimiter = (constVar) ? xRanges[0].Count() - 1 : xRanges[0].Count();
                for (var i = 0; i < delimiter; i++)
                {
                    for (var j = 0; j < xRanges.Count(); j++)
                    {
                        xRanges[j][i] = knownXs[counter];
                        counter += 1;
                    }
                }
            }

            InMemoryRange result;
            if (multipleXranges)
            {
                result = LinestHelper.MultipleRegResult(knownYs, xRanges, constVar, stats, logest);
            }
            else
            {
                result = LinestHelper.LinearRegResult(knownXs, knownYs, constVar, stats, logest);
            }
            return result;
        }

        internal static double[] GetDefaultKnownXs(int count)
        {
            //If no x-values are provided as input, LINEST aranges default values in ascending order
            double[] result = new double[count];
            for (int i = 1; i <= count; i++)
            {
                result[i - 1] = i;
            }
            return result;
        }

        internal static double[] ReverseCoefficientOrder(double[] coefficients, List<double> dropCols, bool constVar, bool logest)
        {
            //This functions puts the coefficients in the same order as excel. The order is reversed (x_n, x_(n - 1), ..., x1)
            //The final order is represented in the array betaCoefficients

            double[] betaCoefficients = new double[coefficients.Count() + dropCols.Count()];
            for (var i = 0; i < dropCols.Count(); i++)
            {
                //If logest we represent collinear columns with 1 instead of 0
                betaCoefficients[betaCoefficients.Count() - ((constVar) ? 1 : 2) - (int)dropCols[i]] = (logest) ? 1d : 0d;
            }

            betaCoefficients[betaCoefficients.Count() - 1] = coefficients[0];
            var count = coefficients.Count() - 1;
            for (var i = 0; i < betaCoefficients.Count() - 1; i++)
            {
                if (!dropCols.Contains(betaCoefficients.Count() - ((constVar) ? 1 : 2) - i))
                {
                    betaCoefficients[i] = coefficients[count];
                    count--;
                }
            }
            
            return betaCoefficients;
        }
    
        internal static void GetSums(double[] coefficients, double[][] xRanges, double[] knownYs, bool logest, bool constVar, out double ssreg, out double ssresid)
        {
            //This function calculates The regression sum of squares and The residual sum of squares
            List<double> estimatedYs = new List<double>(); //This is calculated for each row as y = m1 * x1 + m2 * x2 + ... + mn * xn + intercept (m = coefficient).
            List<double> estimatedErrors = new List<double>(); //This is simply the difference between the observed y-value and the predicted y-value.

            for (var i = 0; i < xRanges.Count(); i++)
            {
                var y = 0d;
                var variables = (constVar) ? coefficients.Count() - 1 : coefficients.Count() - 2;
                for (var k = 0; k < coefficients.Count(); k++)
                {
                    if (logest)
                    {
                        //For LOGEST: Log the coefficients to get rid of EXP
                        y += (k != coefficients.Count() - 1) ? Math.Log(coefficients[k]) * xRanges[i][variables - k] : Math.Log(coefficients[k]);
                    }
                    else
                    {
                        //For LINEST: y = m1 * x1 + m2 * x2 ... mn * xn + b
                        y += (k != coefficients.Count() - 1) ? coefficients[k] * xRanges[i][variables - k] : coefficients[k];
                    }
                }
                estimatedYs.Add(y);
            }

            for (var i = 0; i < estimatedYs.Count; i++)
            {
                var error = knownYs[i] - estimatedYs[i];
                estimatedErrors.Add(error);
            }

            ssresid = (constVar) ? MatrixHelper.DevSq(estimatedErrors, false) : MatrixHelper.DevSq(estimatedErrors, true); //Regression sum of squares
            ssreg = (constVar) ? MatrixHelper.DevSq(estimatedYs, false) : MatrixHelper.DevSq(estimatedYs, true); //Residual sum of squares

        }
        
        internal static double[] GetStandardErrors(double df, double ssresid, int height, int width, List<double> dropCols, double[][] xRangeList)
        {
            //Calculating standard errors of all coefficients below
            var residualMS = (df != 0d) ? ssresid / (height - width + dropCols.Count()) : 0d; //Mean squared of the sum of residual
            var xT = MatrixHelper.TransposeMatrix(xRangeList, height, width - dropCols.Count());
            var xTdotX = MatrixHelper.Multiply(xT, xRangeList);
            var inverseMat = MatrixHelper.Inverse(xTdotX);
            var mIs = (MatrixHelper.GetDeterminant(xTdotX) < 1E-8) ? true : false;

            if (mIs) inverseMat = MatrixHelper.CreateMatrix(inverseMat.Count(), inverseMat[0].Count());

            //Standard errors are derived from the inverse matrix of sum of squares and cross product (SSCP matrix) multiplied with residualMS
            //The standard errors are the squared root of the main diagonal of this matrix.
            var standardErrorMat = MatrixHelper.MatrixMultDouble(inverseMat, residualMS);
            var diagonal = MatrixHelper.MatrixDiagonal(standardErrorMat);
            double[] standardErrorList = new double[diagonal.Count()];
            for (var i = 0; i < standardErrorList.Count(); i++)
            {
                standardErrorList[i] = Math.Sqrt(diagonal[i]);
            }

            //Adjust the standard errors of collinear columns to zero
            double[] standardErrorArray = new double[standardErrorList.Count() + dropCols.Count()];
            var standardIndex = 1;
            if (dropCols.Count() > 0)
            {
                standardErrorArray[0] = standardErrorList[0];
                for (var i = 1; i < standardErrorArray.Count(); i++)
                {
                    if (dropCols.Contains(standardErrorArray.Count() - i - 1))
                    {
                        standardErrorArray[i] = 0d;
                    }
                    else
                    {
                        standardErrorArray[i] = standardErrorList[standardIndex];
                        standardIndex++;
                    }
                }
                return standardErrorArray;
            }
            return standardErrorList;
        }
    }
}
