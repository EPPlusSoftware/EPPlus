using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Xsl;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class MatrixHelper
    {
        public static List<List<double>> TransposedMult(List<List<double>> matrix, double width, double height)
        {
            //This function returns the result of a transposed matrix multiplied by itself.

            List<List<double>> resultMatrix = new List<List<double>>();

            for (int i = 0; i < width; i++)
            {
                List<double> matrixRow = new List<double>();
                for (int j = 0; j < width; j++)
                {
                    var dotSum = 0d;
                    for (int k = 0; k < height; k++)
                    {
                        dotSum += matrix[i][k] * matrix[j][k];
                    }
                    matrixRow.Add(dotSum);
                }
                resultMatrix.Add(matrixRow);
            }

            return resultMatrix;
        }

        internal static double[][] TransposeMatrix(double[][] matrix, int rows, int cols)
        {
            //This function takes a jagged matrix as input, and returns its transpose.
            double[][] transposedMat = CreateMatrix(cols, rows);

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    transposedMat[c][r] = matrix[r][c];
                }
            }
            return transposedMat;
        }
        public static double DevSq(List<double> array, bool meanIsZero)
        {
            //Returns the sum of squares of deviations from a set of datapoints.
            var mean = (!meanIsZero) ? array.Select(x => (double)x).Average() : 0d;
            return array.Aggregate(0d, (val, x) => val += Math.Pow(x - mean, 2));
        }
        public static double[][] MatrixMultDouble(double[][] matrix, double multiplier)
        {
            //Multiplies all elements in a matrix with a single number.
            double[][] resultMat = CreateMatrix(matrix.Count(), matrix[0].Count());
            for (int row = 0; row < matrix.Count(); row++)
            {
                for (int col = 0; col < matrix[0].Count(); col++)
                {
                    resultMat[row][col] = matrix[row][col] * multiplier;
                }
            }
            return resultMat;
        }
        public static double[] MatrixDiagonal(double[][] matrix)
        {
            //Returns the diagonal of a matrix.
            double[] resultArray = new double[matrix.Count()];
            for (int row = 0; row < matrix[0].Count(); row++)
            {
                for (int col = 0; col < matrix.Count(); col++)
                {
                    if (row == col) resultArray[row] = matrix[row][col];
                }
            }
            return resultArray;
        }
        public static List<List<double>> MatrixMult(List<List<double>> matrix1, List<List<double>> matrix2, bool evenDimensions)
        {
            //Calculates the result of multiplying two matrixes
            if (!evenDimensions)
            {
                List<List<double>> resultMatrix = new List<List<double>>();
                for (int i = 0; i < matrix1.Count; i++)
                {
                    List<double> matrixRow = new List<double>();
                    for (int j = 0; j < matrix2[0].Count; j++)
                    {
                        var prodSum = 0d;

                        for (int k = 0; k < matrix1.Count; k++)
                        {
                            prodSum += matrix1[i][k] * matrix2[k][j];
                        }
                        matrixRow.Add(prodSum);
                    }
                    resultMatrix.Add(matrixRow);
                }
                return resultMatrix;
            }
            else
            {
                List<List<double>> resultMatrix = new List<List<double>>();
                return resultMatrix;
            }

        }
        public static List<List<double>> MatrixMultArray(List<List<double>> matrix, List<double> array)
        {
            //Returns the result matrix of a matrix multiplied with an array.
            List<List<double>> resultMatrix = new List<List<double>>();
            for (int i = 0; i < matrix.Count; i++)
            {
                List<double> matrixRow = new List<double>();
                var prodSum = 0d;
                for (int j = 0; j < array.Count; j++)
                {
                    prodSum += matrix[i][j] * array[j];
                }
                matrixRow.Add(prodSum);
                resultMatrix.Add(matrixRow);
            }
            return resultMatrix;
        }
        public static List<List<double>> GetMatrixMinor(List<List<double>> matrix, double i, double j)
        {
            //Returns the minor for a given matrix and entries i:th row and j:th col
            List<List<double>> resultMatrix = new List<List<double>>();
            for (int row = 0; row < matrix.Count; row++)
            {
                if (row == i) continue;

                List<double> matrixRow = new List<double>();
                for (int col = 0; col < matrix[row].Count; col++)
                {
                    if (col == j) continue;

                    matrixRow.Add(matrix[row][col]);
                }
                resultMatrix.Add(matrixRow);
            }
            return resultMatrix;
        }

        internal static double[][] CreateMatrix(int rows, int cols)
        {
            double[][] matrix = new double[rows][];
            for (int i = 0; i < rows; i++)
            {
                matrix[i] = new double[cols];
            }
            return matrix;
        }

        internal static double[][] Multiply(double[][] a, double[][] b)
        {
            int aY = a.Length;
            int aX = a[0].Length;
            int bY = b.Length;
            int bX = b[0].Length;
            if (aX != bY)
            {
                return null;
            }
            double[][] matrix = CreateMatrix(aY, bX);
            for (int i = 0; i < aY; i++)
            {
                for (int j = 0; j < bX; j++)
                {
                    for (int k = 0; k < aX; k++)
                    {
                        matrix[i][j] += a[i][k] * b[k][j];
                    }
                }
            }
            return matrix;
        }

        internal static double[][] GetIdentityMatrix(int size)
        {
            double[][] identity = CreateMatrix(size, size);
            for (int i = 0; i < size; i++)
            {
                identity[i][i] = 1.0d;
            }
            return identity;
        }

        internal static double GetDeterminant(double[][] matrix)
        {
            int[] permutations;
            int rowSwap;
            double[][] lu = Decompose(matrix, out permutations, out rowSwap);
            if (lu == null) return double.NaN;
            double result = rowSwap;
            for (int i = 0; i < lu.Length; ++i)
            {
                result *= lu[i][i];
            }
            return result;
        }

        internal static double GetDeterminant(double[][] lu, int rowSwap)
        {
            if (lu == null) return double.NaN;
            double result = rowSwap;
            for (int i = 0; i < lu.Length; ++i)
            {
                result *= lu[i][i];
            }
            return result;
        }

        internal static double[][] Decompose(double[][] matrix, out int[] permutations, out int rowSwap)
        {
            int rows = matrix.Length;
            int cols = matrix[0].Length;
            double[][] decomposedMatrix = Duplicate(matrix);
            permutations = new int[rows];
            for (int i = 0; i < rows; i++)
            {
                permutations[i] = i;
            }
            rowSwap = 1;
            for (int i = 0; i < rows - 1; i++)
            {
                double maxCols = System.Math.Abs(decomposedMatrix[i][i]);
                int permRow = i;
                for (int j = i + 1; j < rows; j++)
                {
                    if (System.Math.Abs(decomposedMatrix[j][i]) > maxCols)
                    {
                        maxCols = System.Math.Abs(decomposedMatrix[j][i]);
                        permRow = j;
                    }
                }
                if (permRow != i)
                {
                    double[] swapRow = decomposedMatrix[permRow];
                    decomposedMatrix[permRow] = decomposedMatrix[i];
                    decomposedMatrix[i] = swapRow;
                    int swap = permutations[permRow];
                    permutations[permRow] = permutations[i];
                    permutations[i] = swap;
                    rowSwap = -rowSwap;
                }
                if (decomposedMatrix[i][i] == 0.0)
                {
                    int swapRowIndex = -1;
                    for (int row = i + 1; row < rows; row++)
                    {
                        if (decomposedMatrix[row][i] != 0.0)
                            swapRowIndex = row;
                    }
                    if (swapRowIndex == -1) return null;
                    double[] swapRow = decomposedMatrix[swapRowIndex];
                    decomposedMatrix[swapRowIndex] = decomposedMatrix[i];
                    decomposedMatrix[i] = swapRow;
                    int swap = permutations[swapRowIndex];
                    permutations[swapRowIndex] = permutations[i];
                    permutations[i] = swap;
                    rowSwap = -rowSwap;
                }
                for (int j = i + 1; j < rows; j++)
                {
                    decomposedMatrix[j][i] /= decomposedMatrix[i][i];
                    for (int k = i + 1; k < rows; k++)
                    {
                        decomposedMatrix[j][k] -= decomposedMatrix[j][i] * decomposedMatrix[i][k];
                    }
                }
            }
            return decomposedMatrix;
        }

        internal static double[][] Duplicate(double[][] matrix)
        {
            var duplicate = new double[matrix.Length][];
            for (int i = 0; i < matrix.Length; i++)
            {
                var row = matrix[i];
                var newRow = new double[row.Length];
                Array.Copy(row, newRow, row.Length);
                duplicate[i] = newRow;
            }
            return duplicate;
        }

        internal static double[][] Inverse(double[][] matrix)
        {
            double[][] inverse = Duplicate(matrix);
            double[][] lu = Decompose(matrix, out int[] permutations, out int rowSwap);
            if (lu == null) return null;
            double[] unit = new double[matrix.Length];
            for (int i = 0; i < matrix.Length; i++)
            {
                for (int j = 0; j < matrix.Length; j++)
                {
                    if (i == permutations[j])
                    {
                        unit[j] = 1.0;
                    }
                    else
                    {
                        unit[j] = 0.0;
                    }
                }
                double[] element = InverserSolver(lu, unit);
                for (int j = 0; j < matrix.Length; j++)
                {
                    inverse[j][i] = element[j];
                }
            }
            return inverse;
        }

        internal static double[][] Inverse(double[][] lu, int[] permutations, int rowSwap)
        {
            double[][] inverse = Duplicate(lu);
            if (lu == null) return null;
            double[] unit = new double[lu.Length];
            for (int i = 0; i < lu.Length; i++)
            {
                for (int j = 0; j < lu.Length; j++)
                {
                    if (i == permutations[j])
                    {
                        unit[j] = 1.0;
                    }
                    else
                    {
                        unit[j] = 0.0;
                    }
                }
                double[] elements = InverserSolver(lu, unit);
                for (int j = 0; j < lu.Length; j++)
                {
                    inverse[j][i] = elements[j];
                }
            }
            return inverse;
        }

        private static double[] InverserSolver(double[][] luMatrix, double[] unit)
        {
            double[] elements = new double[luMatrix.Length];
            unit.CopyTo(elements, 0);
            for (int i = 1; i < luMatrix.Length; i++)
            {
                double product = elements[i];
                for (int j = 0; j < i; j++)
                {
                    product -= luMatrix[i][j] * elements[j];
                }
                elements[i] = product;
            }
            elements[luMatrix.Length - 1] /= luMatrix[luMatrix.Length - 1][luMatrix.Length - 1];
            for (int i = luMatrix.Length - 2; i >= 0; i--)
            {
                double product = elements[i];
                for (int j = i + 1; j < luMatrix.Length; j++)
                {
                    product -= luMatrix[i][j] * elements[j];
                }
                elements[i] = product / luMatrix[i][i];
            }
            return elements;
        }

        internal static int argMaxAbsolute(double[][] mat)
        {
            //This function finds the index of the largest absolute value in the input matrix

            double maxAbsValue = double.MinValue;
            int maxIndex = -1;
            int flatIndex = 0;

            for (int i = 0; i < mat.Count(); i++)
            {
                for (int j = 0; j < mat[0].Count(); j++)
                {
                    double absValue = Math.Abs(mat[i][j]);
                    if (absValue > maxAbsValue)
                    {
                        maxAbsValue = absValue;
                        maxIndex = flatIndex;
                    }
                    flatIndex += 1;
                }
            }

            return maxIndex;
        }

        internal static double GetM2Norm(double[][] rr)
        {

            //Calculates the 2-norm (euclidean, L2 norm) of the matrix.

            int m = rr.Count();
            int n = rr[0].Count();

            var m1norm = 0d;
            for (int i = 0; i < n; i++)
            {
                //Sum of the max absolute values of all values in column i
                var dd = 0d;
                for (int r = 0; r < rr.Count(); r++)
                {
                    dd += Math.Abs(rr[r][i]);
                    m1norm = Math.Max(m1norm, dd);
                }
            }

            var m8norm = 0d;
            for (int i = 0; i < m; i++)
            {
                //Sum of the max absolute values of all values in row i
                var dd = 0d;
                for (int c = 0; c < rr[0].Count(); c++)
                {
                    dd += Math.Abs(rr[i][c]);
                    m8norm = Math.Max(m8norm, dd);
                }
            }

            return Math.Sqrt(m1norm * m8norm);
        }

        internal static List<double> GaussRank(double[][] xRange, bool constVal)
        {
            //This function takes the input matrix and transforms it to echelon form (every leading coefficient is 1 and is to the right of the leading coefficient on the row above).
            //This is done with complete pivoting to improve numerical stability and identify linearly dependent columns.
            // 
            var xTdotX = MatrixHelper.Multiply(MatrixHelper.TransposeMatrix(xRange, xRange.Count(), xRange[0].Count()), xRange);
            List<double> drop = new List<double>();
            int m = xTdotX.Length;
            int n = xTdotX[0].Length;
            var m2norm = GetM2Norm(xRange);
            var eps = 2.220446049250313E-16;
            var xeps = 1000 * eps;
            int ixr;
            int ixc;

            double[] colOrder = new double[n];
            int count = 0;
            for (int i = 0; i < n; i++)
            {
                colOrder[i] = count;
                count += 1;
            }

            for (int ix0 = 0; ix0 < n; ix0++)
            {
                if (ix0 == 0 && constVal)
                {
                    //If column with 1's has been added, this column is addressed first.
                    ixr = 0;
                    ixc = 0;
                }
                else
                {
                    //Complete pivoting is performed
                    //Pivote element becomes the index with the largest absolute value in each sub matrix
                    double[][] subMatrix = new double[m - ix0][];
                    int rowCount = 0;
                    for (int i = ix0; i < m; i++)
                    {
                        subMatrix[rowCount] = new double[n - ix0];
                        int colCount = 0;
                        for (int j = ix0; j < n; j++)
                        {
                            subMatrix[rowCount][colCount] = xTdotX[i][j];
                            colCount += 1;
                        }
                        rowCount += 1;
                    }
                    int dd = argMaxAbsolute(subMatrix);

                    ixr = dd / (n - ix0);
                    ixc = dd % (n - ix0);
                    ixr += ix0;
                    ixc += ix0;

                    List<double> ddArray = new List<double>();
                    for (int i = 0; i < xTdotX[ixr].Count(); i++)
                    {
                        var tmp = Math.Abs(Math.Abs(xTdotX[ixr][i]) - Math.Abs(xTdotX[ixr][ixc]));
                        if (tmp < 1000 * xeps) ddArray.Add(i);
                    }
                    ixc = (int)ddArray[0];
                    if (ddArray.Count() > 1)
                    {
                        ixc = (int)ddArray[ddArray.Count() - 1];
                    }
                }

                if (Math.Abs(xTdotX[ixr][ixc]) > eps)
                {
                    //row swap
                    for (int i = 0; i < xTdotX[ixr].Count(); ++i)
                    {
                        var tmp = xTdotX[ix0][i];
                        xTdotX[ix0][i] = xTdotX[ixr][i];
                        xTdotX[ixr][i] = tmp;
                    }  
                    //column swap
                    for (int i = 0; i < xTdotX.Count(); i++)
                    {
                        var tmp = xTdotX[i][ix0];
                        xTdotX[i][ix0] = xTdotX[i][ixc];
                        xTdotX[i][ixc] = tmp;
                    }

                    var tmp1 = colOrder[ix0];
                    colOrder[ix0] = colOrder[ixc];
                    colOrder[ixc] = tmp1;

                    //Elimination
                    for (int ix2 = ix0 + 1; ix2 < m; ix2++)
                    {
                        var dd = xTdotX[ix2][ix0] / xTdotX[ix0][ix0];

                        for (int j = 0; j < xTdotX[ix2].Count(); j++)
                        {
                            xTdotX[ix2][j] -= xTdotX[ix0][j] * dd;
                        }
                    }
                  
                }
            }

            //Deciding on collinear columns:
            for (int i = 0; i < n; i++)
            {
                var v1 = xTdotX[i][i];
                var v2 = m2norm;
                var v3 = xeps;
                var v4 = Math.Floor(Math.Abs(v1 / v2) / v3); //.Floor should be correct
                if (v4 == 0)
                {
                    drop.Add(colOrder[i]);
                }
            }

            //Contains column index on what variables should be dropped due to collinearity
            return drop;
        }

        internal static double[][] RemoveColumns(double[][] xRangeList, List<double> dropCols)
        {
            //Removes column indexes in dropCols from the matrix xRangeList
            int height = xRangeList.Length;
            if (height == 0) return xRangeList;

            int width = xRangeList[0].Length;
            HashSet<double> dropColsSet = new HashSet<double>(dropCols);

            double[][] newXRangeList = new double[height][];
            for (int i = 0; i < height; i++)
            {
                List<double> newRow = new List<double>();
                for (int j = 0; j < width; j++)
                {
                    if (!dropColsSet.Contains(j))
                    {
                        newRow.Add(xRangeList[i][j]);
                    }
                }
                newXRangeList[i] = newRow.ToArray();
            }

            return newXRangeList;
        }
    }
}
