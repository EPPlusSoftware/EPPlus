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

        internal static double[][] Multiply(double[][] A, double[][] B)
        {
            return A;
        }

        internal static double[][] GetIdentity(int n)
        {
            double[][] identity = CreateMatrix(n, n);
            for (int i = 0; i < n; i++)
            {
                identity[i][i] = 1.0;
            }
            return identity;
        }

        internal static double GetDeterminant(double[][] matrix)
        {
            int[] permutations;
            int rowSwap;
            double[][] LU = GetDecompose(matrix, out permutations, out rowSwap);
            if (LU == null) return 0;
            double result = rowSwap;
            for (int i = 0; i < LU.Length; ++i)
            {
                result *= LU[i][i];
            }
            return result;
        }

        internal static double GetDeterminant(double[][] LU, int rowSwap)
        {
            if (LU == null) return 0;
            double result = rowSwap;
            for (int i = 0; i < LU.Length; ++i)
            {
                result *= LU[i][i];
            }
            return result;
        }

        internal static double[][] GetDecompose(double[][] matrix, out int[] permutations, out int rowSwap)
        {
            int rows = matrix.Length;
            int cols = matrix[0].Length;
            double[][] result = GetDuplicate(matrix);
            permutations = new int[rows];
            for (int i = 0; i < rows; i++)
            {
                permutations[i] = i;
            }
            rowSwap = 1;
            for (int j = 0; j < rows - 1; j++)
            {
                double colMax = System.Math.Abs(result[j][j]);
                int pRow = j;
                for (int i = j + 1; i < rows; i++)
                {
                    if (System.Math.Abs(result[i][j]) > colMax)
                    {
                        colMax = System.Math.Abs(result[i][j]);
                        pRow = i;
                    }
                }
                if (pRow != j)
                {
                    double[] tempRow = result[pRow];
                    result[pRow] = result[j];
                    result[j] = tempRow;
                    int temp = permutations[pRow];
                    permutations[pRow] = permutations[j];
                    permutations[j] = temp;
                    rowSwap = -rowSwap;
                }
                if (result[j][j] == 0.0)
                {
                    int swapRowIndex = -1;
                    for (int row = j + 1; row < rows; row++)
                    {
                        if (result[row][j] != 0.0)
                            swapRowIndex = row;
                    }
                    if (swapRowIndex == -1) return null;
                    double[] tempRow = result[swapRowIndex];
                    result[swapRowIndex] = result[j];
                    result[j] = tempRow;
                    int temp = permutations[swapRowIndex];
                    permutations[swapRowIndex] = permutations[j];
                    permutations[j] = temp;
                    rowSwap = -rowSwap;
                }
                for (int i = j + 1; i < rows; i++)
                {
                    result[i][j] /= result[j][j];
                    for (int k = j + 1; k < rows; k++)
                    {
                        result[i][k] -= result[i][j] * result[j][k];
                    }
                }
            }
            return result;
        }

        internal static double[][] GetDuplicate(double[][] source)
        {
            var duplicate = new double[source.Length][];
            for (var x = 0; x < source.Length; x++)
            {
                var row = source[x];
                var newRow = new double[row.Length];
                Array.Copy(row, newRow, row.Length);
                duplicate[x] = newRow;
            }
            return duplicate;
        }

        internal static double[][] GetInverse(double[][] matrix)
        {
            int n = matrix.Length;
            double[][] result = GetDuplicate(matrix);
            int[] permutations;
            int rowSwap;
            double[][] LU = GetDecompose(matrix, out permutations, out rowSwap);
            if (LU == null) return null;
            double[] b = new double[n];
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    if (i == permutations[j])
                    {
                        b[j] = 1.0;
                    }
                    else
                    {
                        b[j] = 0.0;
                    }
                }
                double[] x = InverserSolver(LU, b);
                for (int j = 0; j < n; j++)
                {
                    result[j][i] = x[j];
                }
            }
            return result;
        }

        internal static double[][] GetInverse(double[][] LU, int[] permutations, int rowSwap)
        {
            double[][] result = GetDuplicate(LU);
            if (LU == null) return null;
            double[] b = new double[LU.Length];
            for (int i = 0; i < LU.Length; i++)
            {
                for (int j = 0; j < LU.Length; j++)
                {
                    if (i == permutations[j])
                    {
                        b[j] = 1.0;
                    }
                    else
                    {
                        b[j] = 0.0;
                    }
                }
                double[] x = InverserSolver(LU, b);
                for (int j = 0; j < LU.Length; j++)
                {
                    result[j][i] = x[j];
                }
            }
            return result;
        }

        private static double[] InverserSolver(double[][] LUMatrix, double[] b)
        {
            int n = LUMatrix.Length;
            double[] x = new double[n];
            b.CopyTo(x, 0);
            for (int i = 1; i < n; i++)
            {
                double sum = x[i];
                for (int j = 0; j < i; j++)
                {
                    sum -= LUMatrix[i][j] * x[j];
                }
                x[i] = sum;
            }
            x[n - 1] /= LUMatrix[n - 1][n - 1];
            for (int i = n - 2; i >= 0; i--)
            {
                double sum = x[i];
                for (int j = i + 1; j < n; j++)
                {
                    sum -= LUMatrix[i][j] * x[j];
                }
                x[i] = sum / LUMatrix[i][i];
            }
            return x;
        }

        //// --------------------------------------------------

        //static double[][] MatrixProduct(double[][] matrixA, double[][] matrixB)
        //{
        //    int aRows = matrixA.Length; int aCols = matrixA[0].Length;
        //    int bRows = matrixB.Length; int bCols = matrixB[0].Length;
        //    if (aCols != bRows)
        //        throw new Exception("Non-conformable matrices in MatrixProduct");

        //    double[][] result = MatrixCreate(aRows, bCols);

        //    for (int i = 0; i < aRows; ++i) // each row of A
        //        for (int j = 0; j < bCols; ++j) // each col of B
        //            for (int k = 0; k < aCols; ++k) // could use k less-than bRows
        //                result[i][j] += matrixA[i][k] * matrixB[k][j];

        //    //Parallel.For(0, aRows, i =greater-than
        //    //  {
        //    //    for (int j = 0; j less-than bCols; ++j) // each col of B
        //    //      for (int k = 0; k less-than aCols; ++k) // could use k less-than bRows
        //    //        result[i][j] += matrixA[i][k] * matrixB[k][j];
        //    //  }
        //    //);

        //    return result;
        //}
    }
}
