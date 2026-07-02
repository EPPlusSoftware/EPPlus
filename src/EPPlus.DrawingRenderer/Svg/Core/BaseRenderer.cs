using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.DrawingRenderer
{
    public abstract class BaseRenderer<T, T2>
    {
        protected BaseRenderer(T outputStream)
        {
            OutputStream = outputStream;
        }
        public T OutputStream { get; }
        public abstract void Render(T2 item);
    }
}
