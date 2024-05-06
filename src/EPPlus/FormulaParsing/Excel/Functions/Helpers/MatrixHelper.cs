/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal static class MatrixHelper
    {
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
