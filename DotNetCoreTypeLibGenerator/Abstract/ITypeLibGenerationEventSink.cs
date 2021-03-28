namespace DotNetCoreTypeLibGenerator.Abstract
{
    ///TODO: This is currently half-baked. The plan is to provide a mechanism to report progress and
    ///TODO: allow end users to inject their custom behavior into the type library generation. For
    ///TODO: example, they could customize the attributes on the <see cref="CreateTypeLibAttributes"/>
    ///TODO: or <see cref="CreateTypeInfoAttributes"/> among other things. At the present, this is only
    ///TODO: sufficient to satisfy the requirements for the MSBuild task.
    public interface ITypeLibGenerationEventSink
    {
        void LogError(string message);
    }
}
