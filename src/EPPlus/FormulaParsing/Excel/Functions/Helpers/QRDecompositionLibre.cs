using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    internal class QRDecompositionLibre
    {
        internal bool lcl_CalculateQRdecomposition(InMemoryRange pMatA,
                                  Dictionary<int, double> pVecR, int nK, int nN)
        {
            // ScMatrix matrices are zero based, index access (column,row)
            for (int col = 0; col<nK; col++)
            {
                // calculate vector u of the householder transformation
                double fScale = lcl_GetColumnMaximumNorm(pMatA, col, col, nN);
                if (fScale == 0.0)
                {
                    // A is singular
                    return false;
                }
                for (int row = col; row<nN; row++)
                {
                    // pMatA->PutDouble(pMatAGetDouble(col, row)/fScale, col, row);
                    pMatA.SetDouble(row, col, pMatA.GetDouble(row, col) / fScale);
                }


                double fEuclid = lcl_GetColumnEuclideanNorm(pMatA, col, col, nN);
                double fFactor = 1.0 / fEuclid / (fEuclid + Math.Abs(pMatA.GetDouble(col, col)));
                double fSignum = lcl_GetSign(pMatA.GetDouble(col, col));
                //pMatA->PutDouble(pMatA.GetDouble(col, col) + fSignum* fEuclid, col, col);
                pMatA.SetDouble(col, col, pMatA.GetDouble(col, col) + fSignum * fEuclid);
                pVecR[col] = -fSignum* fScale * fEuclid;

                // apply Householder transformation to A
                for (int c=col+1; c<nK; c++)
                {
                    double fSum = lcl_GetColumnSumProduct(pMatA, col, pMatA, c, col, nN);
                    for (int row = col; row<nN; row++)
                    {
                        //pMatA->PutDouble(pMatA->GetDouble(c, row) - fSum * fFactor * pMatA->GetDouble(col, row), c, row);
                        pMatA.SetDouble(row, c, pMatA.GetDouble(row, c) - fSum * fFactor * pMatA.GetDouble(row, col));
                    }
                        
                }
            }
            return true;
        }

        public static double lcl_GetColumnMaximumNorm(InMemoryRange pMatA, int nC, int nR, int nN)
        {
            double fNorm = 0d;
            for (var row = nR; row < nN; row++)
            {
                double fVal = Math.Abs(pMatA.GetDouble(row, nC));
                if (fNorm < fVal)
                {
                    fNorm = fVal;
                }
            }
            return fNorm;
        }

        public static double lcl_GetColumnEuclideanNorm(InMemoryRange pMatA, int nR, int nC, int nN)
        {
            var fNorm = new List<double>();
            for (var row = nR; row < nN; row++)
            {
                fNorm.Add(pMatA.GetDouble(row, nC) * pMatA.GetDouble(row, nC));
            }
            return Math.Sqrt(KahanSum(fNorm));
        }

        public static double KahanSum(List<double> fNorm)
        {
            var sum = 0d;
            var c = 0d;
            for (var i = 0; i < fNorm.Count; i++)
            {
                var y = fNorm[i] - c;
                var t = sum + y;
                c = (t - sum) - y;
                sum = t;
            }
            return sum;
        }

        public static double lcl_GetSign(double fValue)
        {
            var result = (fValue >= 0d) ? 1d : -1d;
            return result;
        }

        public static double lcl_GetColumnSumProduct(InMemoryRange pMatA, int nCa, InMemoryRange pMatB, int nCb, int nR, int nN)
        {
            var fResult = new List<double>();

            for (var row = nR; row < nN; row++)
            {
                fResult.Add(pMatA.GetDouble(row, nCa) * pMatB.GetDouble(row, nCb));
            }
            return KahanSum(fResult);
        }
    }
}
