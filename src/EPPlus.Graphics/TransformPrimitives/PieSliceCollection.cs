using EPPlus.Graphics.Primitives;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlus.Graphics.TransformPrimitives
{
    public class PieSliceCollection : IEnumerable<PieSliceBase>
    {
        const double MaximumDegrees = 360;
        //double totalDegrees = 0;

        double totalValue = 0;

        List<double> _percentages = new List<double>();
        List<PieSliceBase> _slices = new List<PieSliceBase>();

        internal PieSliceCollection()
        {

        }

        public void Add(PieSliceBase slice, double percentage)
        {
            _percentages.Add(percentage);
            _slices.Add(slice);
        }

        public IEnumerator<PieSliceBase> GetEnumerator()
        {
            for (int i = 0; i < _slices.Count; i++)
            {
                yield return _slices[i];
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Clear()
        {
            _slices.Clear();
            _percentages.Clear();
        }

        /// <summary>
        /// Collection containing the richtext objects
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public PieSliceBase this[int Index]
        {
            get
            {
                var item = _slices[Index];
                return item;
            }
        }

    }
}
