using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using NUnit.Framework;
using DotNetCoreTypeLibGenerator.Extensions;

namespace DotNetCoreTypeLibGenerator.Tests.TypeLibraryGenerator
{
    [TestOf(typeof(Structs.SimpleStruct))]
    public class SimpleStructsTest : TypeLibraryTestFixture
    {
        public SimpleStructsTest() : base(typeof(Structs.SimpleStruct)) { }

        [Test]
        public void Verify_Numbers_Of_Struct_Members()
        {
            var guid = typeof(Structs.SimpleStruct).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            typeInfo.UsingTypeAttr(typeAttr =>
            {
                var actual = typeof(Structs.SimpleStruct).GetFields(BindingFlags.Public | BindingFlags.Instance).Length;
                Assert.AreEqual(actual, typeAttr.cVars);
            });
        }

        [Test]
        public void Verify_Names_Of_Struct_Members()
        {
            var guid = typeof(Structs.SimpleStruct).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            typeInfo.UsingTypeAttr(typeAttr =>
            {
                List<string> actualNames = new List<string>();
                for (var memId = 0; memId < typeAttr.cVars; memId++)
                {
                    string[] names = new string[1];
                    typeInfo.GetNames(memId, names, 1, out _);
                    actualNames.Add(names[0]);
                }
                var expectedNames = typeof(Structs.SimpleStruct).GetFields(BindingFlags.Public | BindingFlags.Instance).Select(f => f.Name);
                Assert.AreEqual(expectedNames, actualNames);
            });
        }

        [Test]
        public void Verify_Data_Types_Of_Struct_Members()
        {
            var guid = typeof(Structs.SimpleStruct).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            typeInfo.UsingTypeAttr(typeAttr =>
            {
                List<VarEnum> actualTypes = new List<VarEnum>();
                for (var memId = 0; memId < typeAttr.cVars; memId++)
                {
                    typeInfo.UsingVarDesc(memId, varDesc =>
                    {
                        actualTypes.Add((VarEnum)varDesc.elemdescVar.tdesc.vt);
                    });
                }
                var expectedTypes = typeof(Structs.SimpleStruct).GetFields(BindingFlags.Public | BindingFlags.Instance).Select(f => f.FieldType.GetVarEnum());
                Assert.AreEqual(expectedTypes, actualTypes);
            });
        }
    }
}
