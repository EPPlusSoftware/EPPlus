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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

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
        #region Create Tests
        [TestMethod]
        public void CreateHandlesStandardFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new Sum();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.IsInstanceOfType(functionCompiler, typeof(DefaultCompiler));
        }

        [TestMethod]
        public void CreateHandlesSpecialIfCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new If();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.IsInstanceOfType(functionCompiler, typeof(IfFunctionCompiler));
        }

        [TestMethod]
        public void CreateHandlesSpecialIfErrorCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new IfError();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.IsInstanceOfType(functionCompiler, typeof(IfErrorFunctionCompiler));
        }

        [TestMethod]
        public void CreateHandlesSpecialIfNaCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new IfNa();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.IsInstanceOfType(functionCompiler, typeof(IfNaFunctionCompiler));
        }

        [TestMethod]
        public void CreateHandlesLookupFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new Column();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.IsInstanceOfType(functionCompiler, typeof(LookupFunctionCompiler));
        }

        [TestMethod]
        public void CreateHandlesErrorFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new IsError();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.IsInstanceOfType(functionCompiler, typeof(ErrorHandlingFunctionCompiler));
        }

        [TestMethod]
        public void CreateHandlesCustomFunctionCompiler()
        {
            var functionRepository = FunctionRepository.Create();
            functionRepository.LoadModule(new TestFunctionModule(_context));
            var functionCompilerFactory = new FunctionCompilerFactory(functionRepository, _context);
            var function = new MyFunction();
            var functionCompiler = functionCompilerFactory.Create(function);
            Assert.IsInstanceOfType(functionCompiler, typeof(MyFunctionCompiler));
        }
        #endregion

        #region Nested Classes
        public class TestFunctionModule : FunctionsModule
        {
            public TestFunctionModule(ParsingContext context)
            {
                var myFunction = new MyFunction();
                var customCompiler = new MyFunctionCompiler(myFunction, context);
                base.Functions.Add(MyFunction.Name, myFunction);
                base.CustomCompilers.Add(typeof(MyFunction), customCompiler);
            }
        }

        public class MyFunction : ExcelFunction
        {
            public const string Name = "MyFunction";
            public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
            {
                throw new NotImplementedException();
            }
        }

        public class MyFunctionCompiler : FunctionCompiler
        {
            public MyFunctionCompiler(MyFunction function, ParsingContext context) : base(function, context) { }
            public override CompileResult Compile(IEnumerable<Expression> children)
            {
                throw new NotImplementedException();
            }
        }
        #endregion
    }
}
