using DotNetCoreTypeLibGenerator.Extensions;
using NUnit.Framework;

namespace DotNetCoreTypeLibGenerator.Tests.TypeLibraryGenerator
{
    [TestOf(typeof(SimpleInterface.IUnknowable))]
    public class SimpleInterfaceTest : TypeLibraryTestFixture
    {
        public SimpleInterfaceTest() : base(typeof(SimpleInterface.IUnknowable)) { }

        [Test]
        public void One_TypeInfo_Defined()
        {
            Assert.AreEqual(1, TypeLib.GetTypeInfoCount(), $"There should be only one type info defined in Constants type library, a module. It instead reported {TypeLib.GetTypeInfoCount()}.");
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
    }
}
