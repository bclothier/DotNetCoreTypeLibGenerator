using System.Runtime.InteropServices;

namespace CircularReferencingTypes
{
    [
        ComVisible(true),
        Guid("3019AC7F-7031-4E8A-9379-FA2147217B43")
    ]
    public static class CyclicConstants
    {
        public const int Foo = (int)ThatEnum.ThisBar;
    }

    [
        ComVisible(true),
        Guid("619276E5-E217-4410-83FC-2255AD830D58")
    ]
    public enum ThisEnum
    {
        ThisFoo = CyclicConstants.Foo,
        NotFoo
    }

    [
        ComVisible(true),
        Guid("6E63B41C-B1C6-4C33-BE4E-96BF74D72BBB")
    ]
    public enum ThatEnum
    {
        ThisBar = 42,
        ReallyAFoo = ThisEnum.ThisFoo
    }
}
