using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using NUnit.Framework;
using DotNetCoreTypeLibGenerator.Extensions;

namespace DotNetCoreTypeLibGenerator.Tests.TypeLibraryGenerator
{
    [TestOf(typeof(Constants.Constants))]
    public class ConstantsTest : TypeLibraryTestFixture
    {
        public ConstantsTest() : base(typeof(Constants.Constants)) { }

        [Test]
        public void One_ITypeInfo_Defined()
        {
            Assert.AreEqual(1, TypeLib.GetTypeInfoCount(), $"There should be only one type info defined in Constants type library, a module. It instead reported {TypeLib.GetTypeInfoCount()}.");
        }

        [Test]
        public void Equal_Numbers_Of_Constants()
        {
            TypeLib.GetTypeInfo(0, out var typeInfo);
            typeInfo.UsingTypeAttr(typeAttr =>
            {
                var numOfConsts = typeof(Constants.Constants).GetFields().Count();
                Assert.AreEqual(numOfConsts, typeAttr.cVars, $"There should be {numOfConsts} constants contained within the module, but the representing ITypeInfo reported there were {typeAttr.cVars}.");
            });
        }

        public static IEnumerable<string> Get_Constant_Data_Types()
        {
            return typeof(Constants.Constants)
                    .GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                    .Where(fi => fi.IsLiteral && !fi.IsInitOnly || fi.IsInitOnly && fi.IsStatic)
                    .Select(x => x.Name);
        }

        public static IEnumerable<(string, object)> Get_Constants()
        {
            return typeof(Constants.Constants)
                    .GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                    .Where(fi => fi.IsLiteral && !fi.IsInitOnly || fi.IsInitOnly && fi.IsStatic)
                    .Select(x => (x.Name, x.IsLiteral && !x.IsInitOnly ? x.GetRawConstantValue() : x.GetValue(null)));
        }

        [Test]
        [TestCaseSource(nameof(Get_Constants))]
        public void Verify_Constant_Data_Type((string name, object value) constant)
        {
            var guid = typeof(Constants.Constants).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            var ints = new[] { -1 };
            typeInfo.GetIDsOfNames(new[] { constant.name }, 1, ints);
            typeInfo.UsingVarDesc(ints[0], varDesc =>
            {
                Assert.AreEqual(Enum.Parse(typeof(VarEnum), constant.name.Replace("I_AM_", "")), (VarEnum)varDesc.elemdescVar.tdesc.vt);
            });
        }

        [Test]
        [TestCaseSource(nameof(Get_Constants))]
        public void Verify_Constant_Values((string name, object value) constant)
        {
            var guid = typeof(Constants.Constants).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            var ints = new[] { -1 };
            typeInfo.GetIDsOfNames(new[] { constant.name }, 1, ints);
            typeInfo.UsingVarDesc(ints[0], varDesc =>
            {
                var literalValue = Marshal.GetObjectForNativeVariant(varDesc.desc.lpvarValue);
                var constantValue = constant.value;
                if (constantValue is CurrencyWrapper cy) constantValue = cy.WrappedObject;
                if (constantValue is ErrorWrapper er) constantValue = er.ErrorCode;
                Assert.AreEqual(constantValue, literalValue);
            });
        }
    }
}
