using System.Runtime.InteropServices;

namespace SimpleInterface
{
    [
        ComVisible(true),
        Guid("EEBD79BC-64CE-4BE7-8750-16C9C94A17D8"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)
    ]
    public interface IUnknowable
    {
        int UnknowableNumber { get; set; }
        void DoUnknownThings();
    }
}
