using DotNetCoreTypeLibGenerator.Abstract;
using DotNetCoreTypeLibGenerator.Helpers;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using Vanara.PInvoke;
using static Vanara.PInvoke.OleAut32;
using static DotNetCoreTypeLibGenerator.Helpers.HResultHelper;

namespace DotNetCoreTypeLibGenerator
{
    internal struct CreateTypeLibAttributes
    {
        public string Name { get; set; }
        public ushort MajorVersion { get; set; }
        public ushort MinorVersion { get; set; }
        public LCID LCID { get; set; }
        public LIBFLAGS Flags { get; set; }
        public string DocString { get; set; }
        public uint HelpContext { get; set; }
        public uint HelpStringContext { get; set; }
        public string HelpStringDll { get; set; }
    }

    internal class CreateTypeLibMetadata
    {
        private readonly ITypeLibGenerationEventSink _sink;
        private readonly IReferencedTypeLibrariesProvider _referencedTypeLibProvider;

        public SYSKIND SysKind { get; }
        public Guid Guid { get; }
        public Assembly Assembly { get; }
        public ICreateTypeLib2 CreateTypeLib { get; }
        public ITypeLib2 TypeLib => (ITypeLib2)CreateTypeLib;
        private readonly Dictionary<Guid, CreateTypeInfoMetadata> _typeInfoMetaData;
        public IReadOnlyDictionary<Guid, CreateTypeInfoMetadata> TypeInfoMetaData { get => _typeInfoMetaData; }
        public CreateTypeLibAttributes Attributes { get; }
        public string TlbPath { get; }

        public CreateTypeLibMetadata(SYSKIND sysKind, Guid guid, Assembly assembly, ICreateTypeLib2 createTypeLib, string tlbPath, ITypeLibGenerationEventSink sink, IReferencedTypeLibrariesProvider referencedTypeLibProvider)
        {
            _sink = sink;
            TlbPath = tlbPath;

            SysKind = sysKind;
            Assembly = assembly;
            CreateTypeLib = createTypeLib;
            Guid = guid;

            if (Guid == null || Guid == Guid.Empty)
            {
                throw new ArgumentException($"The assembly {assembly.FullName} does not have a GUID assigned. The type must have a Guid attribute.");
            }

            var name = assembly.GetName();
            var version = name.Version;
            Attributes = new CreateTypeLibAttributes
            {
                Name = AttributeHelper.GetComAliasName(assembly) ?? name.Name,
                MajorVersion = (ushort)version.Major,
                MinorVersion = (ushort)version.Minor,
                LCID = 0,
                Flags = 0
            };

            CreateTypeLib.SetName(Attributes.Name);
            CreateTypeLib.SetGuid(guid);
            CreateTypeLib.SetVersion(Attributes.MajorVersion, Attributes.MinorVersion);
            CreateTypeLib.SetLcid(Attributes.LCID);
            CreateTypeLib.SetLibFlags((uint)Attributes.Flags);

            _typeInfoMetaData = new Dictionary<Guid, CreateTypeInfoMetadata>();
            _referencedTypeLibProvider = referencedTypeLibProvider;
        }

        public void AddType(TYPEKIND typeKind, Type type, ITypeInfo referencedTypeInfo = null)
        {
            SUCCEEDED(CreateTypeLib.CreateTypeInfo(type.Name, typeKind, out var createTypeInfo));
            var createTypeInfoMetadata = new CreateTypeInfoMetadata(typeKind, type, (ICreateTypeInfo2)createTypeInfo, _sink, _referencedTypeLibProvider);
            _typeInfoMetaData.Add(type.GUID, createTypeInfoMetadata);

            // Interfaces must implement a base interface that ultimately derives from the IUnknown or IDispatch
            // interfaces. See for more details:
            // https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-icreatetypeinfo-addimpltype
            switch (typeKind)
            {
                case TYPEKIND.TKIND_DISPATCH:
                    _referencedTypeLibProvider.TryGetReferencedTypeInfo(WellKnown.Iids.IID_DISPATCH, out var dispatchTypeInfo);
                    uint hrefDispatch = 0;
                    SUCCEEDED(createTypeInfo.AddRefTypeInfo(dispatchTypeInfo, hrefDispatch));
                    SUCCEEDED(createTypeInfo.AddImplType(WellKnown.ImplIndexes.BaseInterface, hrefDispatch));
                    if (referencedTypeInfo != null)
                    {
                        uint hrefReferenced = 0;
                        SUCCEEDED(createTypeInfo.AddRefTypeInfo(referencedTypeInfo, hrefReferenced));
                        SUCCEEDED(createTypeInfo.AddImplType(WellKnown.ImplIndexes.DispatchInterface, hrefReferenced));
                    }
                    break;
                case TYPEKIND.TKIND_INTERFACE:
                    _referencedTypeLibProvider.TryGetReferencedTypeInfo(WellKnown.Iids.IID_UNKNOWN, out var unknownTypeInfo);
                    uint hrefUnknown = 0;
                    SUCCEEDED(createTypeInfo.AddRefTypeInfo(unknownTypeInfo, hrefUnknown));
                    SUCCEEDED(createTypeInfo.AddImplType(WellKnown.ImplIndexes.BaseInterface, hrefUnknown));
                    break;
            }
        }

        public void AddTypes(TYPEKIND typeKind, IEnumerable<Type> types)
        {
            foreach (var type in types)
            {
                AddType(typeKind, type);
            }
        }
    }
}
