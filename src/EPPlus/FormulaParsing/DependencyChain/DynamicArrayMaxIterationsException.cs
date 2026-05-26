using System;

namespace OfficeOpenXml.FormulaParsing.DependencyChain
{

    /// <summary>
    /// The maximum number of iterations has been reached when recalculating a dynamic array formula. 
    /// </summary>
    public class DynamicArrayMaxIterationsException : Exception
    {
        /// <summary>
        /// Constructs a new instance of the DynamicArrayMaxIterationsException class.
        /// </summary>
        public DynamicArrayMaxIterationsException()
        {
        }
        /// <summary>
        /// Constructs a new instance of the DynamicArrayMaxIterationsException class with a specified error message.
        /// </summary>
        /// <param name="message">The error message</param>
        public DynamicArrayMaxIterationsException(string message) : base(message)
        {
        }
    }
}
