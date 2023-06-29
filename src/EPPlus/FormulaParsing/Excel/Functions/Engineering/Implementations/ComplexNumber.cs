using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Implementations
{
    internal class ComplexNumber
    {
        public ComplexNumber(double real, double imaginary, string imagSuffix)
        {
            Real = real;
            Imaginary = imaginary;
            ImagSuffix = imagSuffix;
        }

        public double Real { get; }
        public double Imaginary { get; }
        public string ImagSuffix { get; }

        public ComplexNumber GetProduct(ComplexNumber other)
        {
            // (ac - bd) + (ad + bc)
            var real1 = this.Real * other.Real;
            var real2 = this.Imaginary * other.Imaginary;
            var imag1 = this.Real * other.Imaginary;
            var imag2 = this.Imaginary * other.Real;

            var result = new ComplexNumber(real1 - real2, imag1 + imag2, ImagSuffix);
            return result;
        }
    }
}
