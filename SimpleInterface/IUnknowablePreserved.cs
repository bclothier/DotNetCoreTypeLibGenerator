using System.Runtime.InteropServices;

namespace SimpleInterface
{
    [
        ComVisible(true),
        Guid("EEBD79BC-64CE-4BE7-8750-16C9C94A17D9"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)
    ]
    public interface IUnknowablePreserved
    {
        [PreserveSig]
        double DoUnknownThings();
    }
}
