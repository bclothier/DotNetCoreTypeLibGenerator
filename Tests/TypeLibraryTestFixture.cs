using System;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using NUnit.Framework;
using static Vanara.PInvoke.OleAut32;

namespace DotNetCoreTypeLibGenerator.Tests
{
    [SingleThreaded, FixtureLifeCycle(LifeCycle.SingleInstance)]
    public abstract class TypeLibraryTestFixture
    {
        protected readonly Assembly TargetAssembly;
        protected string TlbPath { get; private set; }
        protected bool TlbGenerated { get; private set; }
        protected ITypeLib2 TypeLib { get; private set; }

        public TypeLibraryTestFixture(Type type) : this(type.Assembly) { }

        public TypeLibraryTestFixture(Assembly assembly)
        {
            TargetAssembly = assembly;
        }

        [OneTimeSetUp]
        public void OneTimeSetup()
        {
            var generator = new TypeLibGenerator(TargetAssembly);
            TlbGenerated = generator.GenerateTypeLib(TypeLibGenerator.InferSysKindFromPlatform(), out var tlbPath);            
            TlbPath = tlbPath;

            LoadTypeLib(tlbPath, out var typeLib);
            TypeLib = (ITypeLib2)typeLib;
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            TypeLib = null;
        }
    }
}
