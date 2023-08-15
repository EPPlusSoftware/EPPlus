/*
 * Licensed to the Apache Software Foundation (ASF) under one or more
 * contributor license agreements.  See the NOTICE file distributed with
 * this work for additional information regarding copyright ownership.
 * The ASF licenses this file to You under the Apache License, Version 2.0
 * (the "License"); you may not use this file except in compliance with
 * the License.  You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers
{
    // Calculates the QR-decomposition of given Matrix X. X = QR where Q is an orthogonal matrix and R is an upper triangular matrix.
    // The QR-decomposition is calculated using Householder Reflectors.
    // Parts of the code are ported from http://home.apache.org/~evanward/commons-math-3.6.1-RC1-site/jacoco/org.apache.commons.math3.linear/QRDecomposition.java.html licensed under Apache Software Foundation (ASF)
    /**
     * Calculates the QR-decomposition of a matrix.
     * <p>The QR-decomposition of a matrix A consists of two matrices Q and R
     * that satisfy: A = QR, Q is orthogonal (Q<sup>T</sup>Q = I), and R is
     * upper triangular. If A is m&times;n, Q is m&times;m and R m&times;n.</p>
     * <p>This class compute the decomposition using Householder reflectors.</p>
     * <p>For efficiency purposes, the decomposition in packed form is transposed.
     * This allows inner loop to iterate inside rows, which is much more cache-efficient
     * in Java.</p>
     * <p>This class is based on the class with similar name from the
     * <a href="http://math.nist.gov/javanumerics/jama/">JAMA</a> library, with the
     * following changes:</p>
     * <ul>
     *   <li>a {@link #getQT() getQT} method has been added,</li>
     *   <li>the {@code solve} and {@code isFullRank} methods have been replaced
     *   by a {@link #getSolver() getSolver} method and the equivalent methods
     *   provided by the returned {@link DecompositionSolver}.</li>
     * </ul>
     *
     * @see <a href="http://mathworld.wolfram.com/QRDecomposition.html">MathWorld</a>
     * @see <a href="http://en.wikipedia.org/wiki/QR_decomposition">Wikipedia</a>
     *
     * @since 1.2 (changed to concrete class in 3.0)
     */
    internal class QRDecomposition
    {
        /**
         * A packed TRANSPOSED representation of the QR decomposition.
         * <p>The elements BELOW the diagonal are the elements of the UPPER triangular
         * matrix R, and the rows ABOVE the diagonal are the Householder reflector vectors
         * from which an explicit form of Q can be recomputed if desired.</p>
         */
        private List<List<double>> _qrt;
        /** The diagonal elements of R. */
        private double[] _rDiag;
        /** Cached value of Q. */
        private InMemoryRange _cachedQ;
        /** Cached value of QT. */
        private InMemoryRange _cachedQT;
        /** Cached value of R. */
        private InMemoryRange _cachedR;
        /** Cached value of H. */
        private InMemoryRange _cachedH;
        /** Singularity threshold. */
        private double _threshold;

        /**
         * Calculates the QR-decomposition of the given matrix.
         * The singularity threshold defaults to zero.
         *
         * @param matrix The matrix to decompose.
         *
         * @see #QRDecomposition(InMemoryRange,double)
         */
        public QRDecomposition(InMemoryRange matrix)
            : this(matrix, 0d)
        {
        }

        /**
         * Calculates the QR-decomposition of the given matrix.
         *
         * @param matrix The matrix to decompose.
         * @param threshold Singularity threshold.
         */

        public QRDecomposition(InMemoryRange matrix,
                               double threshold)
        {
            _threshold = threshold;

            int m = matrix.Size.NumberOfRows;
            int n = matrix.Size.NumberOfCols;
            _qrt = matrix.Transpose().ToDoubleMatrix();
            _rDiag = new double[FastMathHelper.Min(m, n)];
            _cachedQ = null;
            _cachedQT = null;
            _cachedR = null;
            _cachedH = null;

            Decompose(_qrt);

        }

        /** Decompose matrix.
         * @param matrix transposed matrix
         * @since 3.2
         */
        protected void Decompose(List<List<double>> matrix)
        {
            for (int minor = 0; minor < FastMathHelper.Min(matrix.Count, matrix[0].Count); minor++)
            {
                performHouseholderReflection(minor, matrix);
            }
        }

        /** Perform Householder reflection for a minor A(minor, minor) of A.
         * @param minor minor index
         * @param matrix transposed matrix
         * @since 3.2
         */
        protected void performHouseholderReflection(int minor, List<List<double>> matrix)
        {

            double[] qrtMinor = matrix[minor].ToArray();

            /*
             * Let x be the first column of the minor, and a^2 = |x|^2.
             * x will be in the positions qr[minor][minor] through qr[m][minor].
             * The first column of the transformed minor will be (a,0,0,..)'
             * The sign of a is chosen to be opposite to the sign of the first
             * component of x. Let's find a:
             */
            double xNormSqr = 0;
            for (int row = minor; row < qrtMinor.Length; row++)
            {
                double c = qrtMinor[row];
                xNormSqr += c * c;
            }
            double a = (qrtMinor[minor] > 0) ? -Math.Sqrt(xNormSqr) : Math.Sqrt(xNormSqr);
            _rDiag[minor] = a;

            if (a != 0.0)
            {

                /*
                 * Calculate the normalized reflection vector v and transform
                 * the first column. We know the norm of v beforehand: v = x-ae
                 * so |v|^2 = <x-ae,x-ae> = <x,x>-2a<x,e>+a^2<e,e> =
                 * a^2+a^2-2a<x,e> = 2a*(a - <x,e>).
                 * Here <x, e> is now qr[minor][minor].
                 * v = x-ae is stored in the column at qr:
                 */
                qrtMinor[minor] -= a; // now |v|^2 = -2a*(qr[minor][minor])

                /*
                 * Transform the rest of the columns of the minor:
                 * They will be transformed by the matrix H = I-2vv'/|v|^2.
                 * If x is a column vector of the minor, then
                 * Hx = (I-2vv'/|v|^2)x = x-2vv'x/|v|^2 = x - 2<x,v>/|v|^2 v.
                 * Therefore the transformation is easily calculated by
                 * subtracting the column vector (2<x,v>/|v|^2)v from x.
                 *
                 * Let 2<x,v>/|v|^2 = alpha. From above we have
                 * |v|^2 = -2a*(qr[minor][minor]), so
                 * alpha = -<x,v>/(a*qr[minor][minor])
                 */
                for (int col = minor + 1; col < matrix.Count; col++)
                {
                    double[] qrtCol = matrix[col].ToArray();
                    double alpha = 0;
                    for (int row = minor; row < qrtCol.Length; row++)
                    {
                        alpha -= qrtCol[row] * qrtMinor[row];
                    }
                    alpha /= a * qrtMinor[minor];

                    // Subtract the column vector alpha*v from x.
                    for (int row = minor; row < qrtCol.Length; row++)
                    {
                        qrtCol[row] -= alpha * qrtMinor[row];
                    }
                }
            }
        }


        /**
         * Returns the matrix R of the decomposition.
         * <p>R is an upper-triangular matrix</p>
         * @return the R matrix
         */
        public InMemoryRange getR()
        {

            if (_cachedR == null)
            {

                // R is supposed to be m x n
                int n = _qrt.Count;
                int m = _qrt[0].Count;
                //double[][] ra = new double[m][n];
                var ir = new InMemoryRange(m, (short)n);
                //var ir = new InMemoryRange(FastMathHelper.Min(n, m), (short)(FastMathHelper.Min(n, m)));
                // copy the diagonal from rDiag and the upper triangle of qr
                for (int row = FastMathHelper.Min(m, n) - 1; row >= 0; row--)
                {
                    //ra[row][row] = _rDiag[row];
                    ir.SetValue(row, row, _rDiag[row]);
                    for (int col = row + 1; col < n; col++)
                    {
                        //ra[row][col] =_qrt[col][row];
                        ir.SetValue(row, col, _qrt[col][row]);
                    }
                }
                //_cachedR = MatrixUtils.createInMemoryRange(ra);
                _cachedR = ir;
            }

            // return the cached matrix
            return _cachedR;
        }

        /**
         * Returns the matrix Q of the decomposition.
         * <p>Q is an orthogonal matrix</p>
         * @return the Q matrix
         */
        public InMemoryRange getQ()
        {
            if (_cachedQ == null)
            {
                _cachedQ = getQT().Transpose();
            }
            return _cachedQ;
        }

        /**
         * Returns the transpose of the matrix Q of the decomposition.
         * <p>Q is an orthogonal matrix</p>
         * @return the transpose of the Q matrix, Q<sup>T</sup>
         */
        public InMemoryRange getQT()
        {
            if (_cachedQT == null)
            {

                // QT is supposed to be m x m
                int n = _qrt.Count;
                //double[][] qta = new double[m][m];
                int m = _qrt[0].Count;
                var ir = new InMemoryRange(m, (short)m); 
                //var ir = new InMemoryRange(m, (short)n);

                /*
                 * Q = Q1 Q2 ... Q_m, so Q is formed by first constructing Q_m and then
                 * applying the Householder transformations Q_(m-1),Q_(m-2),...,Q1 in
                 * succession to the result
                 */
                for (int minor = m - 1; minor >= FastMathHelper.Min(m, n); minor--)
                {
                    //qta[minor][minor] = 1.0d;
                    ir.SetValue(minor, minor, 1d);
                }

                for (int minor = FastMathHelper.Min(m, n) - 1; minor >= 0; minor--)
                {
                    double[] qrtMinor = _qrt[minor].ToArray();
                    //qta[minor][minor] = 1.0d;
                    ir.SetValue(minor, minor, 1d);
                    if (qrtMinor[minor] != 0.0)
                    {
                        for (int col = minor; col < m; col++)
                        {
                            double alpha = 0;
                            for (int row = minor; row < m; row++)
                            {
                                //alpha -= qta[col][row] * qrtMinor[row]; ********************************************************************************************************************************
                                alpha -= ConvertUtil.GetValueDouble(ir.GetValue(row, col)) * qrtMinor[row];
                            }
                            alpha /= _rDiag[minor] * qrtMinor[minor];

                            for (int row = minor; row < m; row++)
                            {
                                //qta[col][row] += -alpha * qrtMinor[row];
                                var currentVal = ConvertUtil.GetValueDouble(ir.GetValue(row, col));
                                ir.SetValue(row, col, currentVal + (-alpha * qrtMinor[row]));
                            }
                        }
                    }
                }
                //_cachedQT = MatrixUtils.createInMemoryRange(qta);
                _cachedQT = ir;
            }

            // return the cached matrix
            return _cachedQT;
        }

        /**
         * Returns the Householder reflector vectors.
         * <p>H is a lower trapezoidal matrix whose columns represent
         * each successive Householder reflector vector. This matrix is used
         * to compute Q.</p>
         * @return a matrix containing the Householder reflector vectors
         */
        public InMemoryRange getH()
        {
            if (_cachedH == null)
            {

                int n = _qrt.Count;
                int m = _qrt[0].Count;
                //double[][] ha = new double[m][n];
                var ir = new InMemoryRange(m, (short)n);
                for (int i = 0; i < m; ++i)
                {
                    for (int j = 0; j < FastMathHelper.Min(i + 1, n); ++j)
                    {
                        //ha[i][j] = qrt[j][i] / -rDiag[j];
                        ir.SetValue(i, j, _qrt[j][i] / -_rDiag[j]);
                    }
                }
                //_cachedH = MatrixUtils.createInMemoryRange(ha);
                _cachedH = ir;
            }

            // return the cached matrix
            return _cachedH;
        }

        /**
         * Get a solver for finding the A &times; X = B solution in least square sense.
         * <p>
         * Least Square sense means a solver can be computed for an overdetermined system,
         * (i.e. a system with more equations than unknowns, which corresponds to a tall A
         * matrix with more rows than columns). In any case, if the matrix is singular
         * within the tolerance set at {@link QRDecomposition#QRDecomposition(InMemoryRange,
         * double) construction}, an error will be triggered when
         * the {@link DecompositionSolver#solve(RealVector) solve} method will be called.
         * </p>
         * @return a solver
         */
        internal Solver getSolver()
        {
            return new Solver(_qrt, _rDiag, _threshold);
        }

        /** Specialized solver. */
        internal class Solver
        {
            /**
             * A packed TRANSPOSED representation of the QR decomposition.
             * <p>The elements BELOW the diagonal are the elements of the UPPER triangular
             * matrix R, and the rows ABOVE the diagonal are the Householder reflector vectors
             * from which an explicit form of Q can be recomputed if desired.</p>
             */
            private List<List<double>> _qrt;
            /** The diagonal elements of R. */
            private double[] _rDiag;
            /** Singularity threshold. */
            private double _threshold;

            /**
             * Build a solver from decomposed matrix.
             *
             * @param qrt Packed TRANSPOSED representation of the QR decomposition.
             * @param rDiag Diagonal elements of R.
             * @param threshold Singularity threshold.
             */
            internal Solver(List<List<double>> qrt,
                            double[] rDiag,
                            double threshold)
            {
                this._qrt = qrt;
                this._rDiag = rDiag;
                this._threshold = threshold;
            }

            /** {@inheritDoc} */
            public bool isNonSingular()
            {
                foreach (var diag in _rDiag)
                {
                    if (Math.Abs(diag) <= _threshold)
                    {
                        return false;
                    }
                }
                return true;
            }

            /** {@inheritDoc} */
            public double[] Solve(double[] b)
            {
                int n = _qrt.Count;
                int m = _qrt[0].Count;
                if (b.Length != m)
                {
                    //throw new DimensionMismatchException(b.getDimension(), m);
                    throw new InvalidOperationException("Invalid dimension in Solve()");
                }
                if (!isNonSingular())
                {
                    //throw new SingularMatrixException();
                    throw new InvalidOperationException("SingularMatrixException!");
                }

                double[] x = new double[n];
                var bCopy = new double[b.Length];
                Array.Copy(b, bCopy, b.Length);
                double[] y = bCopy;

                // apply Householder transforms to solve Q.y = b
                for (int minor = 0; minor < FastMathHelper.Min(m, n); minor++)
                {

                    double[] qrtMinor = _qrt[minor].ToArray();
                    double dotProduct = 0;
                    for (int row = minor; row < m; row++)
                    {
                        dotProduct += y[row] * qrtMinor[row];
                    }
                    dotProduct /= _rDiag[minor] * qrtMinor[minor];

                    for (int row = minor; row < m; row++)
                    {
                        y[row] += dotProduct * qrtMinor[row];
                    }
                }

                // solve triangular system R.x = y
                for (int row = _rDiag.Length - 1; row >= 0; --row)
                {
                    y[row] /= _rDiag[row];
                    double yRow = y[row];
                    double[] qrtRow = _qrt[row].ToArray();
                    x[row] = yRow;
                    for (int i = 0; i < row; i++)
                    {
                        y[i] -= yRow * qrtRow[i];
                    }
                }

                //return new ArrayRealVector(x, false); //RealVector as InMemoryRange(1, x.Length) --> x.Count should just be n, which is equal to _qrt.Count
                return x;
            }

            /** {@inheritDoc} */
            public InMemoryRange SolveMat(InMemoryRange b)
            {
                int n = _qrt.Count;
                int m = _qrt[0].Count;
                if (b.Size.NumberOfRows != m)
                {
                    //throw new DimensionMismatchException(b.getRowDimension(), m);
                    throw new InvalidOperationException("Invalid dimension in Solve()");
                }
                if (!isNonSingular())
                {
                    //throw new SingularMatrixException();
                    throw new InvalidOperationException("SingularMatrixException!");
                }

                int columns = b.Size.NumberOfCols;
                int blockSize = BlockMatrix.BLOCK_SIZE;
                int cBlocks = (columns + blockSize - 1) / blockSize;
                double[][] xBlocks = BlockMatrix.CreateBlockMatrix(n, columns);
                //var y = new double[b.Size.NumberOfRows, blockSize];
                var y = new List<List<double>>();
                double[] alpha = new double[blockSize];

                for (int kBlock = 0; kBlock < cBlocks; ++kBlock)
                {
                    int kStart = kBlock * blockSize;
                    int kEnd = FastMathHelper.Min(kStart + blockSize, columns);
                    int kWidth = kEnd - kStart;

                    // get the right hand side vector
                    //b.copySubMatrix(0, m - 1, kStart, kEnd - 1, y);
                    var arr = ((InMemoryRange)b.GetOffset(0, kStart, m - 1, kEnd - 1)).ToDoubleMatrix();
                    for (var c = 0; c < arr.Count; c++)
                    {
                        List<double> insertRow = new List<double>();
                        for (var r = 0; r < arr[c].Count; r++)
                        {
                            //y[c, r] = arr[c][r];
                            insertRow.Add(arr[c][r]);
                        }
                        y.Add(insertRow);
                    }
                    // apply Householder transforms to solve Q.y = b
                    for (int minor = 0; minor < FastMathHelper.Min(m, n); minor++)
                    {
                        double[] qrtMinor = _qrt[minor].ToArray();
                        double factor = 1.0 / (_rDiag[minor] * qrtMinor[minor]);
                        //Arrays.fill(alpha, 0, kWidth, 0.0);
                        for(var x = 0; x < kWidth; x++)
                        {
                            alpha[x] = 0d;
                        }
                        for (int row = minor; row < m; ++row)
                        {
                            double d = qrtMinor[row];
                            double[] yRow = y[row].ToArray();
                            for (int k = 0; k < kWidth; ++k)
                            {
                                alpha[k] += d * yRow[k];
                            }
                        }
                        for (int k = 0; k < kWidth; ++k)
                        {
                            alpha[k] *= factor;
                        }

                        for (int row = minor; row < m; ++row)
                        {
                            double d = qrtMinor[row];
                            double[] yRow = y[row].ToArray();
                            for (int k = 0; k < kWidth; ++k)
                            {
                                yRow[k] += alpha[k] * d;
                            }
                        }
                    }

                    // solve triangular system R.x = y
                    for (int j = _rDiag.Length - 1; j >= 0; --j)
                    {
                        int jBlock = j / blockSize;
                        int jStart = jBlock * blockSize;
                        double factor = 1.0 / _rDiag[j];
                        double[] yJ = y[j].ToArray();
                        double[] xBlock = xBlocks[jBlock * cBlocks + kBlock];
                        int index = (j - jStart) * kWidth;
                        for (int k = 0; k < kWidth; ++k)
                        {
                            yJ[k] *= factor;
                            xBlock[index++] = yJ[k];
                        }

                        double[] qrtJ = _qrt[j].ToArray();
                        for (int i = 0; i < j; ++i)
                        {
                            double rIJ = qrtJ[i];
                            double[] yI = y[i].ToArray();
                            for (int k = 0; k < kWidth; ++k)
                            {
                                yI[k] -= yJ[k] * rIJ;
                            }
                        }
                    }
                }

                //return new BlockRealMatrix(n, columns, xBlocks, false);
                var ir = new InMemoryRange(xBlocks.Length, (short)xBlocks[0].Length);
                for (var r = 0; r < xBlocks.Length; ++r)
                {
                    for (var c = 0; c < xBlocks[0].Length; c++)
                    {
                        ir.SetValue(r, c, xBlocks[r][c]);
                    }
                }
                return ir;
            }

            /**
             * {@inheritDoc}
             * @throws SingularMatrixException if the decomposed matrix is singular.
             */
            //public InMemoryRange getInverse()
            //{
            //    return Solve(MatrixUtils.createRealIdentityMatrix(_qrt[0].length));
            //}
        }
    }
}

