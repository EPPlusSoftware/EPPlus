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
        public static double DevSq(List<double> array, bool meanIsZero)
        {
            //Returns the sum of squares of deviations from a set of datapoints.
            var mean = (!meanIsZero) ? array.Select(x => (double)x).Average() : 0d;
            return array.Aggregate(0d, (val, x) => val += Math.Pow(x - mean, 2));
        }
        public static List<List<double>> MatrixMultDouble(List<List<double>> matrix, double multiplier)
        {
            //Multiplies all elements in a matrix with a single number.
            List<List<double>> resultMatrix = new List<List<double>>();
            for (var row = 0; row < matrix.Count; row++)
            {
                List<double> matrixRow = new List<double>();
                for (var col = 0; col < matrix[0].Count; col++)
                {
                    matrixRow.Add(matrix[row][col] * multiplier);
                }
                resultMatrix.Add(matrixRow);
            }
            return resultMatrix;
        }
        public static List<double> MatrixDiagonal(List<List<double>> matrix)
        {
            //Returns the diagonal of a matrix.
            List<double> resultArray = new List<double>();
            for (var row = 0; row < matrix[0].Count; row++)
            {
                for (var col = 0; col < matrix.Count; col++)
                {
                    if (row == col) resultArray.Add(matrix[row][col]);
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
    }
}