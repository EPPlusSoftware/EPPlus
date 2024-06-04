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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Text;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class LinestHelper
    {
        public static InMemoryRange CalculateMultipleXRanges(double[] knownYs, double[][] xRangeList, bool constVar, bool stats)
        {
            //if (constVar)
            //{
            //    List<double> onesArray = new List<double>();
            //    for (var i = 0; i < xRangeList[0].Count(); i++)
            //    {
            //        onesArray.Add(1d);
            //    }
            //    xRangeList.Add(onesArray);
            //}
            var dropCols = MatrixHelper.GaussRank(xRangeList, constVar);
            if (constVar)
            {
                for (var i = 0; i < xRangeList.Count(); i++)
                {
                    xRangeList[i][xRangeList[0].Count() - 1] = 1d; //Might have to change this, revise for row and column cases!
                }
            }

            var width = xRangeList[0].Count();
            var height = xRangeList.Count();
            for (var i = 0; i < dropCols.Count(); i++)
            {

            }

            //Add check if all values in a column are the same, that variable is "redundant in that case!
            var multipleRegressionSlopes = GetSlope(xRangeList, knownYs, constVar, stats, out bool matrixIsSingular);
            double[][] nonCollinearX = new double[height][];
            if (matrixIsSingular)
            {
                //var xRangeListCopy = new List<List<double>>(); //Create copy of independentVariables to use in calculations!
                double[][] xRangeListCopy = new double[xRangeList.Count()][];
                //xRangeList.ForEach(x => xRangeListCopy.Add(x));
                for (var r = 0; r < xRangeList.Count(); r++)
                {
                    xRangeListCopy[r] = new double[xRangeList[r].Count()];
                    for (var c = 0; c < xRangeList[0].Count(); c++)
                    {
                        xRangeListCopy[r][c] = xRangeList[r][c];
                    }
                }

                var singleRegressionData = PerformCollinearityCheck(knownYs, xRangeListCopy, constVar);
                var rSquaredValues = singleRegressionData.Item1;
                var coefficients = singleRegressionData.Item2;
                //var new_mat = MatrixHelper.CollinearityTransformer(knownYs, xRangeList, coefficients);
                //var threshold = 1.93294034300795E-06;
                var threshold = 0.05;

                List<double> collinearityColumns = new List<double>();
                for (var i = 0; i < rSquaredValues.Count() - 1; i++) //This can be optimized!
                {
                    if (Math.Abs(rSquaredValues[i] - rSquaredValues[i + 1]) <= threshold) //Test this threshold thoroughly
                    {
                        if (!(collinearityColumns.Contains(i))) collinearityColumns.Add(i);
                        if (!(collinearityColumns.Contains(i + 1))) collinearityColumns.Add(i + 1);
                    }
                    //for (var j = 0; j < rSquaredValues.Count(); j++)
                    //{
                    //    if (i == j) continue;
                    //    if (Math.Abs(rSquaredValues[i] - rSquaredValues[j]) <= threshold) //Test this threshold thoroughly
                    //    {
                    //        if (!(collinearityColumns.Contains(i))) collinearityColumns.Add(i);
                    //        if (!(collinearityColumns.Contains(j))) collinearityColumns.Add(j);
                    //    }
                    //}
                }

                var saveThisValue = coefficients.Min(x => Math.Abs(x));
                var saveThisColumn = 0;
                for (var i = 0; i < coefficients.Count(); i++)
                {
                    if (Math.Abs(coefficients[i]) == saveThisValue) saveThisColumn = i;
                }

                for(var i = 0; i < xRangeListCopy.Count(); i++)
                {
                    nonCollinearX[i] = new double[1];
                    nonCollinearX[i][0] = xRangeListCopy[i][saveThisColumn];
                }

                //for (var i = 0; i < xRangeListCopy.Count(); i++)
                //{
                //    if (collinearityColumns.Contains(i) && i != saveThisColumn)
                //    {
                //        xRangeListCopy[i] = null;
                //    }
                //}
                ////xRangeListCopy.RemoveAll(x => x == null);
                //xRangeListCopy = xRangeListCopy.Where(xArray => xArray != null).ToArray();
                //for (var i = 0; i < xRangeListCopy.Count(); i++)
                //{
                //    xRangeListCopy[i] = xRangeListCopy[i].Where(x => x != null).ToArray();
                //}
                //removedColumns = collinearityColumns.Count - 1;
                //multipleRegressionSlopes = GetSlope(xRangeListCopy, knownYs, constVar, stats, out bool matIsSingular);
                //multipleRegressionSlopes = GetSlope(nonCollinearX, knownYs, constVar, stats, out bool matIsSingular);
                
                //Temporary, remove this solution. Nothing wrong with it but is bad
                var nominator = 0d;
                var denominator = 0d;
                var averageX = 0d;
                for (var i = 0; i < nonCollinearX.Count(); i++)
                {
                    averageX += nonCollinearX[i][0];
                }
                averageX /= nonCollinearX.Count();
                var averageY = knownYs.Average();

                for (var i = 0; i < knownYs.Count(); i++)
                {
                    var y = knownYs[i];
                    var x = nonCollinearX[i][0];

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

                //populate multipleRegressionSlopes with zeros where column was removed due to collinearity
                var size = (constVar) ? xRangeList[0].Count() - 1: xRangeList[0].Count();
                double[][] tmpArray = new double[size + 1][];
                for (var i = 0; i < collinearityColumns.Count(); i++)
                {
                    tmpArray[i] = new double[1];
                    //if (i != saveThisColumn) multipleRegressionSlopes.Insert((int)collinearityColumns[i], insertZero);
                    if (i != saveThisColumn)
                    {
                        tmpArray[i][0] = 0d;
                    }
                    else
                    {
                        tmpArray[i][0] = m;
                    }
                }
                tmpArray[tmpArray.Count() - 1] = new double[1];
                tmpArray[tmpArray.Count() - 1][0] = b;
                multipleRegressionSlopes = tmpArray;
            }
            //if (!constVar)
            //{
            //    double[] zeroIntercept = new double[multipleRegressionSlopes.Count()]; //when const is false, GetSlopes doesnt return an intercept, so we have to add it manually.
            //    zeroIntercept[0] = 0d;
            //    multipleRegressionSlopes[multipleRegressionSlopes.Count()] = zeroIntercept;
            //}

            if (!stats)
            {
                var resultRange = new InMemoryRange(1, (short)(multipleRegressionSlopes.Count()));
                if (constVar)
                {
                    resultRange.SetValue(0, multipleRegressionSlopes.Count() - 1, multipleRegressionSlopes[multipleRegressionSlopes.Count() - 1][0]);
                }
                else
                {
                    resultRange.SetValue(0, multipleRegressionSlopes.Count() - 1, 0d);
                }

                //Linest returns the coefficients in reversed order, so we iterate through the list from the end to get the correct order.
                int pos = 0;
                for (var i = multipleRegressionSlopes.Count() - 2; i >= 0; i--)
                {
                    resultRange.SetValue(0, pos++, multipleRegressionSlopes[i][0]);
                }
                return resultRange;
            }
            else
            {
                var resultRangeStats = new InMemoryRange(5, (short)(multipleRegressionSlopes.Count()));
                if (constVar)
                {
                    resultRangeStats.SetValue(0, multipleRegressionSlopes.Count() - 1, multipleRegressionSlopes[multipleRegressionSlopes.Count() - 1][0]);
                }
                else
                {
                    resultRangeStats.SetValue(0, multipleRegressionSlopes.Count() - 1, 0d);
                }

                //Linest returns the coefficients in reversed order, so we iterate through the list from the end to get the correct order.
                int pos2 = 0;
                for (var i = multipleRegressionSlopes.Count() - 2; i >= 0; i--)
                {
                    resultRangeStats.SetValue(0, pos2++, multipleRegressionSlopes[i][0]);
                }

                List<double> standardErrorSlopes = new List<double>();
                List<double> estimatedYs = new List<double>(); //This is calculated for each row as y = m1 * x1 + m2 * x2 + ... + mn * xn + intercept
                List<double> estimatedErrors = new List<double>();

                for (var i = 0; i < height; i++)
                {
                    var y = 0d;
                    for (var k = 0; k < width; k++) //check here... was mult.count before
                    {
                        y += (k != multipleRegressionSlopes.Count() - 1) ? multipleRegressionSlopes[k][0] * xRangeList[i][k] : multipleRegressionSlopes[k][0];
                    }
                    estimatedYs.Add(y);
                }

                for (var i = 0; i < estimatedYs.Count; i++) //Each error, the error is the difference between the dependent variable and the predicted (estimated) value
                {
                    var error = knownYs[i] - estimatedYs[i];
                    estimatedErrors.Add(error);
                }

                var ssresid = (constVar) ? MatrixHelper.DevSq(estimatedErrors, false) : MatrixHelper.DevSq(estimatedErrors, true);
                var ssreg = (constVar) ? MatrixHelper.DevSq(estimatedYs, false) : MatrixHelper.DevSq(estimatedYs, true);
                var rSquared = ssreg / (ssreg + ssresid);
                var df = Math.Max(height - width, 0d);
                var standardErrorEstimate = (df != 0d) ? Math.Sqrt(ssresid / df) : 0d;
                object fStatistic = 0d;
                if (df != 0)
                {
                    fStatistic = (constVar) ? (ssreg / (width - 1)) / (ssresid / df) : (ssreg / width) / (ssresid / df);
                }
                else
                {
                    fStatistic = ExcelErrorValue.Create(eErrorType.Num);
                }

                //Calculating standard errors of all coefficients below
                var residualMS = (df != 0d) ? ssresid / (height - width) : 0d; //Mean squared of the sum of residual
                var xT = MatrixHelper.TransposeMatrix(xRangeList, height, width);
                var xTdotX = MatrixHelper.Multiply(xT, xRangeList);
                //var inverseMat = GetInverse(xTdotX, out bool mIs);
                var inverseMat = MatrixHelper.Inverse(xTdotX);
                var mIs = (MatrixHelper.GetDeterminant(xTdotX) < 1E-8) ? true : false; //Have not tested this threshold

                if (mIs)
                {
                    for (var i = 0; i < inverseMat.Count(); i++)
                    {
                        for (var j = 0; j < inverseMat[0].Count(); j++)
                        {
                            inverseMat[i][j] = 0d;
                        }
                    }
                }

                var standardErrorMat = MatrixHelper.MatrixMultDouble(inverseMat, residualMS);
                var diagonal = MatrixHelper.MatrixDiagonal(standardErrorMat);
                double[] standardErrorList = new double[diagonal.Count()];
                for (var i = 0; i < standardErrorList.Count(); i++)
                {
                    standardErrorList[i] = Math.Sqrt(diagonal[i]);
                }
                //var standardErrorList = diagonal.Select(x => Math.Sqrt(x)).ToList(); //Standard errors are derived from the inverse matrix of sum of squares and cross product (SSCP matrix) multiplied with residualMS
                                                                                     //The standard errors are the squared root of the main diagonal of this matrix.

                if (constVar)
                {
                    resultRangeStats.SetValue(1, xRangeList[0].Count() - 1, standardErrorList[standardErrorList.Count() - 1]);
                }
                else
                {
                    resultRangeStats.SetValue(1, xRangeList[0].Count(), ExcelErrorValue.Create(eErrorType.NA));
                }

                int pos3 = 0;
                for (var i = (constVar) ? standardErrorList.Count() - 2 : standardErrorList.Count() - 1; i >= 0; i--)
                {
                    resultRangeStats.SetValue(1, pos3++, standardErrorList[i]);
                }

                if (constVar)
                {
                    for (var col = 2; col < width; col++) //wrong in this loop
                    {
                        for (var row = 2; row < 5; row++)
                        {
                            resultRangeStats.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                        }
                    }
                }
                else
                {
                    for (var col = 2; col < width + 1; col++) //wrong in this loop
                    {
                        for (var row = 2; row < 5; row++)
                        {
                            resultRangeStats.SetValue(row, col, ExcelErrorValue.Create(eErrorType.NA));
                        }
                    }
                }

                resultRangeStats.SetValue(2, 0, rSquared);
                resultRangeStats.SetValue(2, 1, standardErrorEstimate);
                resultRangeStats.SetValue(3, 0, fStatistic);
                resultRangeStats.SetValue(3, 1, df);
                resultRangeStats.SetValue(4, 0, ssreg);
                resultRangeStats.SetValue(4, 1, ssresid);
                return resultRangeStats;
            }
        }

        private static double[][] GetSlope(double[][] xValues, double[] yValues, bool constVar, bool stats, out bool matrixIsSingular)
        {
            var width = xValues[0].Count();
            var height = xValues.Count();
            //var xTdotX = MatrixHelper.TransposedMult(xValues, width, height); //Multiply transpose of X with X, denoted X'X
            var xT = MatrixHelper.TransposeMatrix(xValues, height, width);
            var xTdotX = MatrixHelper.Multiply(xT, xValues);
            var myInverse = MatrixHelper.Inverse(xTdotX);
            //var myInverse = GetInverse(xTdotX, out bool mIs); //Inverse of transpose of X multiplied by X, denoted as (X'X)^-1
            //matrixIsSingular = mIs; matrixIsSingular is an out bool for this function.
            //var dotProduct = MatrixHelper.MatrixMult(myInverse, xValues, false);
            var dotProduct = MatrixHelper.Multiply(myInverse, xT);
            //var b = MatrixHelper.MatrixMultArray(dotProduct, yValues); // b = (X'X)^-1 * X' * Y
            double[][] yValuesJagged = yValues.Select(yVal => new double[] { yVal }).ToArray();
            var b = MatrixHelper.Multiply(dotProduct, yValuesJagged);

            //new code
            matrixIsSingular = (MatrixHelper.GetDeterminant(xTdotX) < 1E-8) ? true : false; //Have not tested this threshold

            if (!constVar)
            {
                double[][] extendedB = new double[b.Count() + 1][];
                for (var i = 0; i < b.Count(); i++)
                {
                    extendedB[i] = new double[1];
                    extendedB[i][0] = b[i][0];
                }
                extendedB[extendedB.Count() - 1] = new double[1];
                extendedB[extendedB.Count() - 1][0] = 0d;
                return extendedB;
            }
            return b;
        }
        private static double GetDeterminant(List<List<double>> matrix)
        {
            if (matrix.Count == 2)
            {
                var determinantTwoByTwo = matrix[0][0] * matrix[1][1] - matrix[0][1] * matrix[1][0]; //If matrix is 2x2, determinant is easily derived from this determinant formula
                return determinantTwoByTwo;
            }
            var determinant = 0d;
            for (int col = 0; col < matrix[0].Count; col++)
            {
                determinant += Math.Pow(-1, col) * matrix[0][col] * GetDeterminant(MatrixHelper.GetMatrixMinor(matrix, 0, col));
            }
            return determinant;
        }
        private static List<List<double>> GetInverse(List<List<double>> mat, out bool matrixIsSingular)
        {
            matrixIsSingular = false;
            var determinant = GetDeterminant(mat);
            //if (Math.Abs(determinant) < 1e-6) matrixIsSingular = true; //check if this threshold holds.
            if (determinant < 1E-8) matrixIsSingular = true; //Have not tested this threshold
            if (mat.Count == 2)
            {
                List<List<double>> twoByTwoMat = new List<List<double>>(); //Same here, special case for a 2x2 matrix
                List<double> row1 = new List<double>();
                List<double> row2 = new List<double>();
                row1.Add(mat[1][1] / determinant);
                row1.Add(-1 * mat[0][1] / determinant);
                row2.Add(-1 * mat[1][0] / determinant);
                row2.Add(mat[0][0] / determinant);
                twoByTwoMat.Add(row1);
                twoByTwoMat.Add(row2);
                return twoByTwoMat;
            }

            List<List<double>> coMat = new List<List<double>>();
            for (var row = 0; row < mat.Count; row++)
            {
                List<double> coRow = new List<double>();
                for (var col = 0; col < mat[row].Count; col++)
                {
                    var minor = MatrixHelper.GetMatrixMinor(mat, row, col);
                    coRow.Add(Math.Pow(-1, row + col) * GetDeterminant(minor));
                }
                coMat.Add(coRow);
            }

            List<List<double>> finalMat = new List<List<double>>();
            for (var row = 0; row < coMat[0].Count; row++)
            {
                List<double> finalRow = new List<double>();
                for (var col = 0; col < coMat.Count; col++)
                {
                    finalRow.Add(coMat[col][row] / determinant);
                }
                finalMat.Add(finalRow);
            }

            return finalMat;
        }

        public static TupleClass<double[], double[]> PerformCollinearityCheck(double[] dependentVariable, double[][] independentVariables, bool constVar)
        {
            //List<double> rSquaredValues = new List<double>();
            //List<double> coefficients = new List<double>();
            //List<double> ones = new List<double>();
            //for (var i = 0; i < independentVariables[0].Count(); i++)
            //{
            //    ones.Add(1d);
            //}
            var size = (constVar) ? independentVariables[0].Count() - 1 : independentVariables[0].Count();
            double[] rSquaredValues = new double[size];
            double[] coefficients = new double[size];
            for (var i = 0; i < size; i++)
            {
                //bool OnesArray = independentVariables[i].TrueForAll(x => x == 1d);
                //if (!OnesArray)
                //if (constVar == true && i == independentVariables.Count() - 1)
                //{
                //    continue;
                //}
                double[] independentVariable = new double[independentVariables.Count()];
                for (var j = 0; j < independentVariables.Count(); j++)
                {
                    independentVariable[j] = independentVariables[j][i];
                }

                var singleRegression = CalculateResult(dependentVariable, independentVariable, constVar, true);
                var rSquared = singleRegression.GetValue(2, 0);
                var coefficient = singleRegression.GetValue(0, 0);
                //rSquaredValues.Add((double)rSquared);
                //coefficients.Add((double)coefficient);
                rSquaredValues[i] = (double)rSquared;
                coefficients[i] = (double)coefficient;
            }

            return new TupleClass<double[], double[]>(rSquaredValues, coefficients);
        }

        public static InMemoryRange CalculateResult(double[] knownYs, double[] knownXs, bool constVar, bool stats)
        {
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

            if (stats)
            {
                for (var i = 0; i < knownXs.Count(); i++)
                {
                    var x = knownXs[i];
                    var y = knownYs[i];
                    var estimatedY = m * x + b;

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
                    df = knownXs.Count() - 2; //Need to review this
                    v1 = knownXs.Count() - df - 1;
                    v2 = df;
                    fStatistics = (ssr / v1) / (yDiff / v2);
                }
                else
                {
                    df = knownXs.Count() - 1; //Need to review this
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
    }
}