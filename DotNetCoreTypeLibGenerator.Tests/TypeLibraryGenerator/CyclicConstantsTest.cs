using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using DotNetCoreTypeLibGenerator.Extensions;
using NUnit.Framework;

namespace DotNetCoreTypeLibGenerator.Tests.TypeLibraryGenerator
{
    [TestOf(typeof(CircularReferencingTypes.CyclicConstants))]
    public class CyclicConstantsTest : TypeLibraryTestFixture
    {
        public CyclicConstantsTest() : base(typeof(CircularReferencingTypes.CyclicConstants)) { }

        [Test]
        public void Verify_Constant_Values()
        {
            var guid = typeof(CircularReferencingTypes.CyclicConstants).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            var ints = new[] { -1 };
            typeInfo.GetIDsOfNames(new[] { "Foo" }, 1, ints);
            typeInfo.UsingVarDesc(0, varDesc =>
            {
                var actual = (int)Marshal.GetObjectForNativeVariant(varDesc.desc.lpvarValue);
                Assert.AreEqual(CircularReferencingTypes.CyclicConstants.Foo, actual);
            });
        }
    }
}
