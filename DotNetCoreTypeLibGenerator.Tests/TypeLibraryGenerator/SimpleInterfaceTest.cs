using DotNetCoreTypeLibGenerator.Extensions;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace DotNetCoreTypeLibGenerator.Tests.TypeLibraryGenerator
{
    [TestOf(typeof(SimpleInterface.IUnknowable))]
    public class SimpleInterfaceTest : TypeLibraryTestFixture
    {
        public SimpleInterfaceTest() : base(typeof(SimpleInterface.IUnknowable)) { }

        [Test]
        public void One_TypeInfo_Defined()
        {
            Assert.AreEqual(2, TypeLib.GetTypeInfoCount(), $"There should be only one type info defined in Constants type library, a module. It instead reported {TypeLib.GetTypeInfoCount()}.");
        }

        [Test]
        public void Three_FuncDescs_Defined()
        {
            var guid = typeof(SimpleInterface.IUnknowable).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            typeInfo.UsingTypeAttr(attr => {
                Assert.AreEqual(3, attr.cFuncs);
            });
        }

        [Test]
        public void Unpreserved_Signatures_Returns_HRESULT()
        {
            var guid = typeof(SimpleInterface.IUnknowable).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);
            
            var names = new string[] { nameof(SimpleInterface.IUnknowable.DoUnknownThings) };
            var memIds = new int[names.Length];
            var count = names.Length;

            typeInfo.GetIDsOfNames(names, count, memIds);

            typeInfo.UsingFuncDesc(memIds[0], funcDesc => {
                Assert.AreEqual((short)VarEnum.VT_HRESULT, funcDesc.elemdescFunc.tdesc.vt);
            });
        }

        [Test]
        public void Preserved_Signatures_Returns_HRESULT()
        {
            var guid = typeof(SimpleInterface.IUnknowablePreserved).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);

            var names = new string[] { nameof(SimpleInterface.IUnknowablePreserved.DoUnknownThings) };
            var memIds = new int[names.Length];
            var count = names.Length;

            typeInfo.GetIDsOfNames(names, count, memIds);

            typeInfo.UsingFuncDesc(memIds[0], funcDesc => {
                Assert.AreEqual((short)VarEnum.VT_R8, funcDesc.elemdescFunc.tdesc.vt);
            });
        }

        [Test]
        public void Properties_Returns_HRESULT()
        {
            var guid = typeof(SimpleInterface.IUnknowable).GUID;
            TypeLib.GetTypeInfoOfGuid(ref guid, out var typeInfo);

            var names = new string[] { nameof(SimpleInterface.IUnknowable.UnknowableNumber) };
            var memIds = new int[names.Length];
            var count = names.Length;

            typeInfo.GetIDsOfNames(names, count, memIds);

            typeInfo.UsingFuncDesc(memIds[0], funcDesc => {
                Assert.AreEqual((short)VarEnum.VT_HRESULT, funcDesc.elemdescFunc.tdesc.vt);
            });
        }
    }
}
