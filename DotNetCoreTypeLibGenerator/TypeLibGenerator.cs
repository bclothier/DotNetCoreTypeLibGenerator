using DotNetCoreTypeLibGenerator.Abstract;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

using SYSKIND = System.Runtime.InteropServices.ComTypes.SYSKIND;

namespace DotNetCoreTypeLibGenerator
{
    public class TypeLibGenerator 
    {
        private readonly ITypeLibGenerationEventSink _sink;
        private readonly Assembly _assembly;
        private readonly string _tlbPath;

        public static Assembly GetAssemblyFromPath(string assemblyPath)
        {
            var localPath = new Uri(assemblyPath).LocalPath;
            if (!File.Exists(localPath))
            {
                throw new FileNotFoundException($"Cannot find the assembly: {assemblyPath}");
            }
            return Assembly.LoadFrom(localPath);
        }

        public static string GetDefaultTlbPath(Assembly assembly)
        {
            var localPath = new Uri(assembly.CodeBase).LocalPath;
            return Path.Combine(Path.GetDirectoryName(localPath), Path.GetFileNameWithoutExtension(localPath) + ".tlb");
        }

        public static SYSKIND InferSysKindFromPlatform()
        {
            return Marshal.SizeOf<IntPtr>() == 8 ? SYSKIND.SYS_WIN64 : SYSKIND.SYS_WIN32;
        }

        public TypeLibGenerator(Assembly assembly) : 
            this(assembly, null) { }

        public TypeLibGenerator(Assembly assembly, ITypeLibGenerationEventSink sink) : 
            this(GetDefaultTlbPath(assembly), assembly, sink) { }

        public TypeLibGenerator(string tlbPath, Assembly assembly, ITypeLibGenerationEventSink sink)
        {
            _sink = sink;
            _assembly = assembly;
            _tlbPath = tlbPath;
        }

        public bool GenerateTypeLib(out string tlbPath)
        {
            return GenerateTypeLib(InferSysKindFromPlatform(), out tlbPath);
        }

        public bool GenerateTypeLib(SYSKIND sysKind, out string tlbPath)
        {
            var factory = new CreateTypeLibraryFactory(_sink);
            factory.Create(sysKind, _assembly, _tlbPath);
            tlbPath = _tlbPath;
            return true;
        }
    }
}
