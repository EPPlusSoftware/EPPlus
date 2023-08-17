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
        public static InMemoryRange CalculateMultipleXRanges(List<double> knownYs, List<List<double>> xRangeList, bool constVar, bool stats)
        {
            if (constVar)
            {
                //var range = InMemoryRange.FromDoubleMatrix(xRangeList);
                //var qr2 = new QRDecompositionLibre();
                //var h2 = new Dictionary<int, double>();
                //qr2.lcl_CalculateQRdecomposition(range, h2, range.Size.NumberOfCols, range.Size.NumberOfRows);
                //var qr = new QRDecomposition(range);
                //var q = qr.getQ();
                //var r = qr.getR();
                //var qT = q.Transpose();
                //var rT = r.Transpose();
                //var qrSolver = qr.getSolver();
                //double[] yArray = new double[knownYs.Count];
                //for (var i = 0; i < knownYs.Count; i++) yArray[i] = knownYs[i];
                //var resultArray = qrSolver.Solve(yArray);
                //var decomposedMat = qrSolver.SolveMat(range);
                //var q = qr.getQ();
                //var r = qr.getR();
                //var blockMatrix = qrSolver.SolveMat();

                //var correlationMatrix = GetCorrelationMatrix(xRangeList);
                //var removeTheseColumns = GetRedundantSetIndex(correlationMatrix);
                //List<double> varianceInflationFactors = new List<double>();

                List<double> onesArray = new List<double>();
                for (var i = 0; i < xRangeList[0].Count; i++)
                {
                    onesArray.Add(1d);
                }
                xRangeList.Add(onesArray);
            }

            var width = xRangeList.Count;
            var height = xRangeList[0].Count;

            var multipleRegressionSlopes = GetSlope(xRangeList, knownYs, constVar, stats, out bool matrixIsSingular);
            if (matrixIsSingular)
            {
                if (!constVar) width += 1;
                if (stats)
                {
                    var numMat = new InMemoryRange(5, (short)(width));
                    for (var row = 0; row < 5; row++)
                    {
                        for (var col = 0; col < width; col++)
                        {
                            numMat.SetValue(row, col, ExcelErrorValue.Create(eErrorType.Num));
                        }
                    }

                    return numMat;
                }
                else
                {
                    var numVec = new InMemoryRange(1, (short)(width));
                    for (var col = 0; col < width; col++)
                    {
                        numVec.SetValue(0, col, ExcelErrorValue.Create(eErrorType.Num));
                    }

                    return numVec;
                }
            }
            if (!constVar)
            {
                List<double> zeroIntercept = new List<double>(); //when const is false, GetSlopes doesnt return an intercept, so we have to add it manually.
                zeroIntercept.Add(0d);
                multipleRegressionSlopes.Add(zeroIntercept);
            }

            if (!stats)
            {
                var resultRange = new InMemoryRange(1, (short)(multipleRegressionSlopes.Count));
                if (constVar)
                {
                    resultRange.SetValue(0, multipleRegressionSlopes.Count - 1, multipleRegressionSlopes[multipleRegressionSlopes.Count - 1][0]);
                }
                else
                {
                    resultRange.SetValue(0, multipleRegressionSlopes.Count - 1, 0d);
                }

                //Linest returns the coefficients in reversed order, so we iterate through the list from the end to get the correct order.
                int pos = 0;
                for (var i = multipleRegressionSlopes.Count - 2; i >= 0; i--)
                {
                    resultRange.SetValue(0, pos++, multipleRegressionSlopes[i][0]);
                }
                return resultRange;
            }
            else
            {
                var resultRangeStats = new InMemoryRange(5, (short)(multipleRegressionSlopes.Count));
                if (constVar)
                {
                    resultRangeStats.SetValue(0, multipleRegressionSlopes.Count - 1, multipleRegressionSlopes[multipleRegressionSlopes.Count - 1][0]);
                }
                else
                {
                    resultRangeStats.SetValue(0, multipleRegressionSlopes.Count - 1, 0d);
                }

                //Linest returns the coefficients in reversed order, so we iterate through the list from the end to get the correct order.
                int pos2 = 0;
                for (var i = multipleRegressionSlopes.Count - 2; i >= 0; i--)
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
                        y += (k != multipleRegressionSlopes.Count) ? multipleRegressionSlopes[k][0] * xRangeList[k][i] : multipleRegressionSlopes[k][0];
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
                var df = height - width;
                var standardErrorEstimate = (df != 0d) ? Math.Sqrt(ssresid / df) : 0d;
                var fStatistic = 0d;
                if (df != 0)
                {
                    fStatistic = (constVar) ? (ssreg / (width - 1)) / (ssresid / df) : (ssreg / width) / (ssresid / df);
                }

                //Calculating standard errors of all coefficients below
                var residualMS = (df != 0d) ? ssresid / (height - width) : 0d; //Mean squared of the sum of residual
                var xTdotX = MatrixHelper.TransposedMult(xRangeList, width, height);
                var inverseMat = GetInverse(xTdotX, out bool mIs);
                var standardErrorMat = MatrixHelper.MatrixMultDouble(inverseMat, residualMS);
                var diagonal = MatrixHelper.MatrixDiagonal(standardErrorMat);
                var standardErrorList = diagonal.Select(x => Math.Sqrt(x)).ToList(); //Standard errors are derived from the inverse matrix of sum of squares and cross product (SSCP matrix) multiplied with residualMS
                                                                                     //The standard errors are the squared root of the main diagonal of this matrix.

                if (constVar)
                {
                    resultRangeStats.SetValue(1, xRangeList.Count - 1, standardErrorList[standardErrorList.Count - 1]);
                }
                else
                {
                    resultRangeStats.SetValue(1, xRangeList.Count, ExcelErrorValue.Create(eErrorType.NA));
                }

                int pos3 = 0;
                for (var i = (constVar) ? standardErrorList.Count - 2 : standardErrorList.Count - 1; i >= 0; i--)
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

        private static List<List<double>> GetSlope(List<List<double>> xValues, List<double> yValues, bool constVar, bool stats, out bool matrixIsSingular)
        {
            var width = xValues.Count;
            var height = xValues[0].Count;
            var xTdotX = MatrixHelper.TransposedMult(xValues, width, height); //Multiply transpose of X with X, denoted X'X
            var myInverse = GetInverse(xTdotX, out bool mIs); //Inverse of transpose of X multiplied by X, denoted as (X'X)^-1
            matrixIsSingular = mIs;
            //if (matrixIsSingular) CalculateResult(yValues, xValues[0], constVar, stats); //*************************************************************************************************
            var dotProduct = MatrixHelper.MatrixMult(myInverse, xValues, false);
            var b = MatrixHelper.MatrixMultArray(dotProduct, yValues);
            return b;
        }
        private static double GetDeterminant(List<List<double>> matrix)
        {
            if (matrix.Count == 2)
            {
                var determinantTwoByTwo = matrix[0][0] * matrix[1][1] - matrix[0][1] * matrix[1][0]; //If matrix is 2x2, determinant is easily derived from this determinant formula
                //if (determinantTwoByTwo == 0d && finalCalc) ;
                //if (determinantTwoByTwo == 0d) determinantTwoByTwo = -2.38964E-12;
                return determinantTwoByTwo;
            }
            var determinant = 0d;
            for (int col = 0; col < matrix[0].Count; col++)
            {
                determinant += Math.Pow(-1, col) * matrix[0][col] * GetDeterminant(MatrixHelper.GetMatrixMinor(matrix, 0, col));
            }
            //if (determinant == 0d && finalCalc == true); //if matrix is singular, we set determinant to near zero.
            //if (determinant == 0d) determinant = -2.38964E-12;
            return determinant;
        }
        private static List<List<double>> GetInverse(List<List<double>> mat, out bool matrixIsSingular)
        {
            matrixIsSingular = false;
            var determinant = GetDeterminant(mat);
            if (determinant == 0d) matrixIsSingular = true;
            if (mat.Count == 2 || matrixIsSingular)
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

        public static InMemoryRange CalculateResult(List<double> knownYs, List<double> knownXs, bool constVar, bool stats)
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

            for (var i = 0; i < knownYs.Count; i++)
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

            var m = nominator / denominator;
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

                var errorVariance = yDiff / (knownXs.Count - 2);
                if (!constVar) errorVariance = yDiff / (knownXs.Count() - 1);

                var standardErrorM = (constVar) ? Math.Sqrt(1d / (knownXs.Count - 2d) * estimatedDiff / xDiff) :
                                                  Math.Sqrt(1d / (knownXs.Count - 1d) * estimatedDiff / xDiff);

                object standardErrorB = Math.Sqrt(errorVariance) * Math.Sqrt(1d / knownXs.Count() + Math.Pow(averageX, 2) / xDiff);
                if (!constVar) standardErrorB = ExcelErrorValue.Create(eErrorType.NA);

                var rSquared = ssr / sst;
                var standardErrorEstimateY = (!constVar) ? SEHelper.GetStandardError(knownXs, knownYs, true) :
                                                          SEHelper.GetStandardError(knownXs, knownYs, false);
                var ssreg = ssr;
                var ssresid = (constVar) ? yDiff : (sst - ssr);

                if (constVar)
                {
                    df = knownXs.Count - 2; //Need to review this
                    v1 = knownXs.Count - df - 1;
                    v2 = df;
                    fStatistics = (ssr / v1) / (yDiff / v2);
                }
                else
                {
                    df = knownXs.Count - 1; //Need to review this
                    v1 = knownXs.Count - df;
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
