using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using NUnit.Framework;
using Moq;
using Microsoft.Build.Utilities;
using Microsoft.Build.Framework;
using DotNetCoreTypeLibGenerator.BuildTask;
using static Vanara.PInvoke.OleAut32;

using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;

namespace DotNetCoreTypeLibGenerator.Tests.BuildTask
{
    public class BuildTaskTests
    {
        [Test]
        public void BuildTypeLib()
        {
            var mockEngine = new Mock<FakeBuildEngine>
            {
                CallBase = true
            };
            var engine = mockEngine.Object;

            var generator = new GenerateTypeLibTask()
            {
                TargetAssembly = AssemblyPath(typeof(EmptyTypeLib.EmptyEnum)),
                BuildEngine = engine
            };

            var result = generator.Execute();
            Assert.IsTrue(result, $"Execute method returned unexpected result: {string.Join(Environment.NewLine, engine.LogErrorEvents.Select(e => e.Message))}");
            mockEngine.Verify(x => x.LogErrorEvent(It.IsAny<BuildErrorEventArgs>()), Times.Never);
        }

        private static string AssemblyPath(Type type) => Assembly.GetAssembly(type).CodeBase;
    }
}