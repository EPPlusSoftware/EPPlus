using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class MatrixHelper
    {
        public static List<List<double>> TransposedMult(List<List<double>> matrix, double width, double height)
        {
            //This function returns the result of a transposed matrix multiplied by itself.

            List<List<double>> resultMatrix = new List<List<double>>();

            for (var i = 0; i < width; i++)
            {
                List<double> matrixRow = new List<double>();
                for (var j = 0; j < width; j++)
                {
                    var dotSum = 0d;
                    for (var k = 0; k < height; k++)
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
            double[][] transposedMat = CreateMatrix(rows, cols);

            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    transposedMat[r][c] = matrix[c][r];
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
            for (var row = 0; row < matrix.Count(); row++)
            {
                for (var col = 0; col < matrix[0].Count(); col++)
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
            for (var row = 0; row < matrix[0].Count(); row++)
            {
                for (var col = 0; col < matrix.Count(); col++)
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
                for (var i = 0; i < matrix1.Count; i++)
                {
                    List<double> matrixRow = new List<double>();
                    for (var j = 0; j < matrix2[0].Count; j++)
                    {
                        var prodSum = 0d;

                        for (var k = 0; k < matrix1.Count; k++)
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
            for (var i = 0; i < matrix.Count; i++)
            {
                List<double> matrixRow = new List<double>();
                var prodSum = 0d;
                for (var j = 0; j < array.Count; j++)
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
    }
}

        //Adrians MatrixHelper below ***********************************************************
        //public static List<List<double>> Multiply(List<List<double>> a, List<List<double>> b)
        //{
        //    int aY = a.Count;
        //    int aX = a[0].Count;
        //    int bY = b.Count;
        //    int bX = b[0].Count;
        //    if (aX != bY)
        //    {
        //        return null;
        //    }
        //    List<List<double>> matrix = new List<List<double>>();
        //    for (int i = 0; i < aY; i++)
        //    {
        //        for (int j = 0; j < bX; j++)
        //        {
        //            for (int k = 0; k < aX; k++)
        //            {
        //                matrix[i][j] += a[i][k] * b[k][j];
        //            }
        //        }
        //    }
        //    return matrix;
        //}

        //public static List<List<double>> GetIdentityMatrix(int size)
        //{
        //    List<List<double>> identity = new List<List<double>>();
        //    for (int i = 0; i < size; i++)
        //    {
        //        identity[i][i] = 1.0d;
        //    }
        //    return identity;
        //}

        //public static double GetDeterminant(List<List<double>> matrix)
        //{
        //    List<int> permutations;
        //    int rowSwap;
        //    List<List<double>> lu = Decompose(matrix, out permutations, out rowSwap);
        //    if (lu == null) return double.NaN;
        //    double result = rowSwap;
        //    for (int i = 0; i < lu.Count; ++i)
        //    {
        //        result *= lu[i][i];
        //    }
        //    return result;
        //}

        //public static double GetDeterminant(List<List<double>> lu, int rowSwap)
        //{
        //    if (lu == null) return double.NaN;
        //    double result = rowSwap;
        //    for (int i = 0; i < lu.Count; ++i)
        //    {
        //        result *= lu[i][i];
        //    }
        //    return result;
        //}

        //public static List<List<double>> Decompose(List<List<double>> matrix, out List<int> permutations, out int rowSwap)
        //{
        //    int rows = matrix.Count;
        //    int cols = matrix[0].Count;
        //    List<List<double>> decomposedMatrix = Duplicate(matrix);
        //    permutations = new List<int>();
        //    for (int i = 0; i < rows; i++)
        //    {
        //        permutations[i] = i;
        //    }
        //    rowSwap = 1;
        //    for (int i = 0; i < rows - 1; i++)
        //    {
        //        double maxCols = System.Math.Abs(decomposedMatrix[i][i]);
        //        int permRow = i;
        //        for (int j = i + 1; j < rows; j++)
        //        {
        //            if (System.Math.Abs(decomposedMatrix[j][i]) > maxCols)
        //            {
        //                maxCols = System.Math.Abs(decomposedMatrix[j][i]);
        //                permRow = j;
        //            }
        //        }
        //        if (permRow != i)
        //        {
        //            List<double> swapRow = decomposedMatrix[permRow];
        //            decomposedMatrix[permRow] = decomposedMatrix[i];
        //            decomposedMatrix[i] = swapRow;
        //            int swap = permutations[permRow];
        //            permutations[permRow] = permutations[i];
        //            permutations[i] = swap;
        //            rowSwap = -rowSwap;
        //        }
        //        if (decomposedMatrix[i][i] == 0.0)
        //        {
        //            int swapRowIndex = -1;
        //            for (int row = i + 1; row < rows; row++)
        //            {
        //                if (decomposedMatrix[row][i] != 0.0)
        //                    swapRowIndex = row;
        //            }
        //            if (swapRowIndex == -1) return null;
        //            List<double> swapRow = decomposedMatrix[swapRowIndex];
        //            decomposedMatrix[swapRowIndex] = decomposedMatrix[i];
        //            decomposedMatrix[i] = swapRow;
        //            int swap = permutations[swapRowIndex];
        //            permutations[swapRowIndex] = permutations[i];
        //            permutations[i] = swap;
        //            rowSwap = -rowSwap;
        //        }
        //        for (int j = i + 1; j < rows; j++)
        //        {
        //            decomposedMatrix[j][i] /= decomposedMatrix[i][i];
        //            for (int k = i + 1; k < rows; k++)
        //            {
        //                decomposedMatrix[j][k] -= decomposedMatrix[j][i] * decomposedMatrix[i][k];
        //            }
        //        }
        //    }
        //    return decomposedMatrix;
        //}

        //public static List<List<double>> Duplicate(List<List<double>> matrix)
        //{
        //    return matrix.Select(row => new List<double>(row)).ToList();
        //}

        //public static List<List<double>> Inverse(List<List<double>> matrix)
        //{
        //    List<List<double>> inverse = Duplicate(matrix);
        //    List<List<double>> lu = Decompose(matrix, out List<int> permutations, out int rowSwap);
        //    if (lu == null) return null;
        //    List<double> unit = new List<double>();
        //    for (int i = 0; i < matrix.Count; i++)
        //    {
        //        for (int j = 0; j < matrix.Count; j++)
        //        {
        //            if (i == permutations[j])
        //            {
        //                unit[j] = 1.0;
        //            }
        //            else
        //            {
        //                unit[j] = 0.0;
        //            }
        //        }
        //        List<double> element = InverserSolver(lu, unit);
        //        for (int j = 0; j < matrix.Count; j++)
        //        {
        //            inverse[j][i] = element[j];
        //        }
        //    }
        //    return inverse;
        //}

        //public static List<List<double>> Inverse(List<List<double>> lu, List<int> permutations, int rowSwap)
        //{
        //    List<List<double>> inverse = Duplicate(lu);
        //    if (lu == null) return null;
        //    List<double> unit = new List<double>();
        //    for (int i = 0; i < lu.Count; i++)
        //    {
        //        for (int j = 0; j < lu.Count; j++)
        //        {
        //            if (i == permutations[j])
        //            {
        //                unit[j] = 1.0;
        //            }
        //            else
        //            {
        //                unit[j] = 0.0;
        //            }
        //        }
        //        List<double> elements = InverserSolver(lu, unit);
        //        for (int j = 0; j < lu.Count; j++)
        //        {
        //            inverse[j][i] = elements[j];
        //        }
        //    }
        //    return inverse;
        //}

        //private static List<double> InverserSolver(List<List<double>> luMatrix, List<double> unit)
        //{
        //    double[] elementsArray = new double[unit.Count];
        //    unit.CopyTo(elementsArray, 0);
        //    List<double> elements = new List<double>(elementsArray);
        //    for (int i = 1; i < luMatrix.Count; i++)
        //    {
        //        double product = elements[i];
        //        for (int j = 0; j < i; j++)
        //        {
        //            product -= luMatrix[i][j] * elements[j];
        //        }
        //        elements[i] = product;
        //    }
        //    elements[luMatrix.Count - 1] /= luMatrix[luMatrix.Count - 1][luMatrix.Count - 1];
        //    for (int i = luMatrix.Count - 2; i >= 0; i--)
        //    {
        //        double product = elements[i];
        //        for (int j = i + 1; j < luMatrix.Count; j++)
        //        {
        //            product -= luMatrix[i][j] * elements[j];
        //        }
        //        elements[i] = product / luMatrix[i][i];
        //    }
        //    return elements;
        //}
