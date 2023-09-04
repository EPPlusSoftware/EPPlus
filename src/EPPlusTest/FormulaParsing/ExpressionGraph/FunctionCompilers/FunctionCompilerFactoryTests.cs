/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;

namespace EPPlusTest.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    [TestClass]
    public class FunctionCompilerFactoryTests
    {
        private ParsingContext _context;

        [TestInitialize]
        public void Initialize()
        {
            _context = ParsingContext.Create();
        }

        [TestMethod]
        public void CreateHandlesStandardFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository);
            var function = new SumSubtotal();
            var functionCompiler = functionCompilerFactory.Create(function, ParsingContext.Create());
            Assert.IsInstanceOfType(functionCompiler, typeof(DefaultCompiler));
        }

        [TestMethod]
        public void CreateHandleCustomArrayCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository);
            var function = new Year();
            var functionCompiler = functionCompilerFactory.Create(function, ParsingContext.Create());
            Assert.IsInstanceOfType(functionCompiler, typeof(CustomArrayBehaviourCompiler));
        }

        #region special compiler test

        public class MyFunction : ExcelFunction
        {
            public override int ArgumentMinLength => 1;

            public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
            {
                return CreateResult(2, DataType.Integer);
            }
        }

        internal class MyFunctionCompiler : DefaultCompiler
        {
            public MyFunctionCompiler(ExcelFunction function) : base(function)
            {
                _function = function;
            }

            private readonly ExcelFunction _function;

            public override CompileResult Compile(IEnumerable<CompileResult> children, ParsingContext context)
            {
                return base.Compile(children, context);
            }
        }

        internal class MyModule : IFunctionModule
        {
            public IDictionary<string, ExcelFunction> Functions => new Dictionary<string, ExcelFunction>()
            {
                { "MyFunction", new MyFunction() }
            };

            public IDictionary<Type, FunctionCompiler> CustomCompilers => new Dictionary<Type, FunctionCompiler>()
            {
                { typeof(MyFunction), new MyFunctionCompiler(new MyFunction()) }
            };
        }

        #endregion
    }
}
