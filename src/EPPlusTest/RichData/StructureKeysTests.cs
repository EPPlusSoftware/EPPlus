using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

// Same namespace and usings as StructureKeys.cs, so StructureTypes and RichValueDataType
// resolve exactly as they do in the library.
namespace EPPlusTests.RichData
{
    /// <summary>
    /// Regression tests for StructureKeys, which previously filled a static dictionary lazily
    /// without locking. Concurrent first use from separate packages corrupted it.
    /// The class has no shared state and is safe to run in parallel with other tests.
    /// </summary>
    [TestClass]
    public class StructureKeysTests
    {
        private class KeyExpectation
        {
            public string Structure;
            public string Name;
            public RichValueDataType DataType;
        }

        /// <summary>
        /// Mirrors the registrations in StructureKeys. Errors.Busy is not registered in the library.
        /// </summary>
        private static List<KeyExpectation> GetExpectedKeys()
        {
            var result = new List<KeyExpectation>();
            AddKeys(result, StructureTypes.Error, StructureKeys.Errors.Propagated);
            AddKeys(result, StructureTypes.Error, StructureKeys.Errors.Field);
            AddKeys(result, StructureTypes.Error, StructureKeys.Errors.Spill);
            AddKeys(result, StructureTypes.Error, StructureKeys.Errors.WithSubType);
            AddKeys(result, StructureTypes.LocalImage, StructureKeys.LocalImage.Image);
            AddKeys(result, StructureTypes.WebImage, StructureKeys.WebImage.Image);
            return result;
        }

        private static void AddKeys(List<KeyExpectation> result, string structure, IEnumerable<ExcelRichValueStructureKey> keys)
        {
            foreach (var key in keys)
            {
                result.Add(new KeyExpectation { Structure = structure, Name = key.Name, DataType = key.DataType });
            }
        }

        /// <summary>
        /// Deterministic check that the lookup table is complete and correct.
        /// Guards against mistakes when the table construction is changed.
        /// </summary>
        [TestMethod]
        public void GetKeyDataType_ShouldReturnTheRegisteredTypeForAllKeys()
        {
            foreach (var expected in GetExpectedKeys())
            {
                var actual = StructureKeys.GetKeyDataType(expected.Structure, expected.Name);
                Assert.IsTrue(actual.HasValue, $"No data type for {expected.Structure}.{expected.Name}");
                Assert.AreEqual(expected.DataType, actual.Value, $"Wrong data type for {expected.Structure}.{expected.Name}");
            }

            Assert.IsNull(StructureKeys.GetKeyDataType(StructureTypes.LocalImage, "NoSuchKey"));
            Assert.IsNull(StructureKeys.GetKeyDataType("NoSuchStructure", "NoSuchKey"));
        }

        /// <summary>
        /// Deterministic regression guard. StructureKeys is shared by all packages in the process,
        /// so all its static state must be assigned once by the type initializer and never replaced
        /// or filled lazily on first use. Fails for the original implementation, where the lookup
        /// dictionary was a non-readonly static field filled when Count == 0.
        /// Note: readonly does not stop someone from mutating the dictionary contents later; it guards
        /// against the specific lazy pattern that caused the bug.
        /// </summary>
        [TestMethod]
        public void StructureKeys_ShouldOnlyHaveReadOnlyStaticFields()
        {
            var mutableFields = new List<string>();
            CollectMutableStaticFields(typeof(StructureKeys), mutableFields);
            Assert.AreEqual(0, mutableFields.Count, "Mutable static fields: " + string.Join(", ", mutableFields.ToArray()));
        }

        private static void CollectMutableStaticFields(Type type, List<string> result)
        {
            const BindingFlags fieldFlags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly;
            foreach (var field in type.GetFields(fieldFlags))
            {
                // Skip compiler-generated fields, e.g. cached lambda delegates.
                if (field.Name.StartsWith("<") || field.IsDefined(typeof(CompilerGeneratedAttribute), false))
                {
                    continue;
                }
                if (!field.IsInitOnly && !field.IsLiteral)
                {
                    result.Add(type.Name + "." + field.Name);
                }
            }
            foreach (var nested in type.GetNestedTypes(BindingFlags.Public | BindingFlags.NonPublic))
            {
                // Skip compiler-generated types, e.g. the lambda cache class <>c.
                if (nested.Name.StartsWith("<") || nested.IsDefined(typeof(CompilerGeneratedAttribute), false))
                {
                    continue;
                }
                CollectMutableStaticFields(nested, result);
            }
        }

        /// <summary>
        /// End-to-end check: many independent packages set, save and reload cell pictures in parallel.
        /// Covers process-wide state in the rich data / cell picture code. It cannot reliably catch a
        /// one-time initialization race, since the state is usually initialized by earlier tests.
        /// </summary>
        [TestMethod]
        public void SetCellPicturesInParallel_ShouldNotInterfere()
        {
            // Load the resources once, outside the parallel loop.
            var pictures = new[] { Resources.Png2ByteArray, Resources.Png3ByteArray };
            var errors = new ConcurrentQueue<Exception>();
            var options = new ParallelOptions { MaxDegreeOfParallelism = Math.Max(4, Environment.ProcessorCount * 2) };

            Parallel.For(0, 50, options, i =>
            {
                try
                {
                    using (var package = new ExcelPackage())
                    {
                        var ws = package.Workbook.Worksheets.Add("Sheet1");
                        ws.Cells["A1"].Picture.Set(pictures[i % 2]);
                        ws.Cells["A2"].Picture.Set(pictures[(i + 1) % 2]);

                        using (var ms = new MemoryStream())
                        {
                            package.SaveAs(ms);
                            ms.Position = 0;
                            using (var reloaded = new ExcelPackage(ms))
                            {
                                var reloadedSheet = reloaded.Workbook.Worksheets[0];
                                if (!reloadedSheet.Cells["A1"].Picture.Exists || !reloadedSheet.Cells["A2"].Picture.Exists)
                                {
                                    throw new InvalidOperationException($"Cell picture missing after reload in iteration {i}.");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    errors.Enqueue(ex);
                }
            });

            Assert.AreEqual(0, errors.Count, string.Join(Environment.NewLine + Environment.NewLine, errors.Take(3).Select(x => x.ToString()).ToArray()));
        }
    }
}