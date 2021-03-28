using DotNetCoreTypeLibGenerator.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

namespace DotNetCoreTypeLibGenerator
{
    internal class ReferencedTypeLibrariesProvider : IReferencedTypeLibrariesProvider
    {
        private List<ITypeLib> _referencedTypeLibraries = new List<ITypeLib>();
        private Dictionary<Guid, ITypeInfo> _cachedTypeInfos = new Dictionary<Guid, ITypeInfo>();

        public void AddReferencedTypeLibrary(ITypeLib typeLib)
        {
            AddReferencedTypeLibrary(typeLib, new Guid[] { });
        }

        public void AddReferencedTypeLibrary(ITypeLib typeLib, IEnumerable<string> preCacheGuids)
        {
            AddReferencedTypeLibrary(typeLib, preCacheGuids.Select(g => new Guid(g)));
        }

        public void AddReferencedTypeLibrary(ITypeLib typeLib, IEnumerable<Guid> preCacheGuids)
        {
            _referencedTypeLibraries.Add(typeLib);
            foreach (var guid in preCacheGuids)
            {
                var g = guid;
                typeLib.GetTypeInfoOfGuid(ref g, out var typeInfo);
                if (typeInfo != null)
                {
                    _cachedTypeInfos.Add(guid, typeInfo);
                }
            }
        }

        public bool TryGetReferencedTypeInfo(Guid guid, out ITypeInfo typeInfo)
        {
            if (_cachedTypeInfos.TryGetValue(guid, out typeInfo))
            {
                return true;
            }

            foreach (var typeLib in _referencedTypeLibraries)
            {
                typeLib.GetTypeInfoOfGuid(ref guid, out typeInfo);
                if (typeInfo != null)
                {
                    return true;
                }
            }

            typeInfo = null;
            return false;
        }
    }
}
