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
            var ac = this.Real * other.Real;
            var bd = this.Imaginary * other.Imaginary;
            var ad = this.Real * other.Imaginary;
            var bc = this.Imaginary * other.Real;

            var result = new ComplexNumber(ac - bd, ad + bc, ImagSuffix);
            return result;
        }
    }
}
