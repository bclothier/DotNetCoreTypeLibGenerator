using DotNetCoreTypeLibGenerator.Abstract;
using Microsoft.Build.Framework;
using Microsoft.Build.Utilities;

namespace DotNetCoreTypeLibGenerator.BuildTask
{
    public class GenerateTypeLibTask : Task, ITypeLibGenerationEventSink
    {
        [Required]
        public string TargetAssembly { get; set; }

        public override bool Execute()
        {
            var generator = new TypeLibGenerator(TypeLibGenerator.GetAssemblyFromPath(TargetAssembly));
            return generator.GenerateTypeLib(TypeLibGenerator.InferSysKindFromPlatform(), out _);
        }

        public void LogError(string message)
        {
            Log.LogError(message);
        }
    }
}
