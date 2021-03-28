using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.ComTypes;

namespace DotNetCoreTypeLibGenerator.Abstract
{
    internal interface IReferencedTypeLibrariesProvider
    {
        void AddReferencedTypeLibrary(ITypeLib typeLib);
        void AddReferencedTypeLibrary(ITypeLib typeLib, IEnumerable<Guid> preCacheGuids);
        void AddReferencedTypeLibrary(ITypeLib typeLib, IEnumerable<string> preCacheGuids);
        bool TryGetReferencedTypeInfo(Guid guid, out ITypeInfo typeInfo);
    }
}