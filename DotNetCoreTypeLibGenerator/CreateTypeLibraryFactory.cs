using DotNetCoreTypeLibGenerator.Abstract;
using DotNetCoreTypeLibGenerator.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static Vanara.PInvoke.OleAut32;
using static DotNetCoreTypeLibGenerator.Helpers.HResultHelper;

namespace DotNetCoreTypeLibGenerator
{
    /// <summary>
    /// The factory encapsulates the process of discovering and creating <see cref="ICreateTypeInfo"/> for
    /// each COM-visible types within the given <see cref="Assembly"/>. Because it is possible for a type
    /// to contain a cyclic reference to either itself or to another type that refers back to the original type,
    /// we must resolve the types in 2 passes. The first pass simply enumerates the COM-visible types and generate
    /// the accompanying <see cref="ICreateTypeInfo"/>. In the second pass we start the actual construction of the 
    /// members represented by that type. In this way, we allow the <see cref="CreateTypeInfoMetadata"/> to 
    /// encapsulate the process of describing a member and control the assignment of the MEMBERID. That enables us 
    /// to proceed with describing each types without having to check whether the referenced types/members has been 
    /// described already. For resolving the references to other types, the <see cref="ICreateTypeInfo"/> can be 
    /// cast into a <see cref="ITypeInfo"/> to be then used in describing the referencing member.
    /// </summary>
    internal class CreateTypeLibraryFactory
    {
        private readonly ITypeLibGenerationEventSink _sink;
        private readonly ReferencedTypeLibrariesProvider _referencedTypeLibProvider;
        private CreateTypeLibMetadata _createTypeLibMetadata;
        private IEnumerable<IGrouping<Type, FieldInfo>> _moduleConstants;
        private IEnumerable<Type> _modules;
        private IEnumerable<Type> _enums;
        private IEnumerable<Type> _records;
        private IEnumerable<Type> _interfaces;
        private IEnumerable<Type> _classes;

        public CreateTypeLibraryFactory() : this(null) { }

        public CreateTypeLibraryFactory(ITypeLibGenerationEventSink sink)
        {
            _sink = sink;
            _referencedTypeLibProvider = new ReferencedTypeLibrariesProvider();

            // Interfaces must always derive from the IUnknown or IDispatch. Therefore, we must always
            // include a reference to the stdole32 type library to allow any ICreateTypeInfo describing
            // a TKIND_INTERFACE to implement either or both. 
            SUCCEEDED(LoadTypeLib("stdole32.tlb", out var oleTypeLib));
            _referencedTypeLibProvider.AddReferencedTypeLibrary(oleTypeLib, 
                new[] { WellKnown.Iids.IID_UNKNOWN, WellKnown.Iids.IID_DISPATCH });
        }

        public CreateTypeLibMetadata Create(SYSKIND sysKind, Assembly assembly, string tlbPath)
        {
            _createTypeLibMetadata = CreateTypeLibMetadataForAssembly(sysKind, assembly, tlbPath);
            if (_createTypeLibMetadata == null)
            {
                return null;
            }

            EnumerateAndCreateTypeInfoForTypes();
            DescribeTypeInfos();

            _createTypeLibMetadata.CreateTypeLib.SaveAllChanges();
            return _createTypeLibMetadata;
        }

        private CreateTypeLibMetadata CreateTypeLibMetadataForAssembly(SYSKIND sysKind, Assembly assembly, string tlbPath)
        {
            CreateTypeLib2(sysKind, tlbPath, out var createTypeLib2);

            var guid = AttributeHelper.GetGuid(assembly);
            if (guid == null)
            {
                _sink?.LogError($"The assembly must have a Guid attribute defined. It was not found on the assembly: {assembly.FullName}");
                return null;
            }

            return new CreateTypeLibMetadata(sysKind, guid, assembly, createTypeLib2, tlbPath, _sink, _referencedTypeLibProvider);
        }

        private void EnumerateAndCreateTypeInfoForTypes()
        {
            var allVisibleTypes = _createTypeLibMetadata.Assembly.GetTypes().Where(t => AttributeHelper.GetComVisible(t));
            _moduleConstants = allVisibleTypes.Where(
                    // Consider only public static classes for containing COM-visible constants
                    t => t.IsPublic && !t.IsNestedPublic && t.IsClass && t.IsAbstract && t.IsSealed
                ).SelectMany(
                    t => t.GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy)
                        // Either consts or static readonly fields
                        .Where(fi => (fi.IsLiteral && !fi.IsInitOnly) || (fi.IsInitOnly && fi.IsStatic))
                ).GroupBy(x => x.DeclaringType);
            _modules = _moduleConstants.Where(x => x.Any()).Select(x => x.Key).Distinct();
            _createTypeLibMetadata.AddTypes(TYPEKIND.TKIND_MODULE, _modules);

            _enums = allVisibleTypes.Where(t => t.IsEnum);
            _createTypeLibMetadata.AddTypes(TYPEKIND.TKIND_ENUM, _enums);

            _records = allVisibleTypes.Where(t => t.IsValueType && !t.IsPrimitive && !t.Namespace.StartsWith(nameof(System)) && !t.IsEnum && t.IsPublic);
            _createTypeLibMetadata.AddTypes(TYPEKIND.TKIND_RECORD, _records);

            _interfaces = allVisibleTypes.Where(t => t.IsInterface);
            _createTypeLibMetadata.AddTypes(TYPEKIND.TKIND_INTERFACE, _interfaces);

            _classes = allVisibleTypes.Where(t => t.IsClass && !t.IsAbstract);
            _createTypeLibMetadata.AddTypes(TYPEKIND.TKIND_COCLASS, _classes);
        }

        private void DescribeTypeInfos()
        {
            foreach (var module in _modules)
            {
                // All constants must be contained within a module. MIDL allows us to define constants
                // outside a module but that will not be then visible in the created type library. 
                // To simplify things, we assume that any COM-visible constants will exist in a COM-visible
                // public static .NET classes and we will pretend that this static class is a "module". See: 
                // https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/82f9465b-ae46-474e-87ff-e65e9751affb
                // https://docs.microsoft.com/en-us/windows/win32/midl/module
                // TODO: Handle entry points for external DLLs which must be also within a module as well
                var createTypeInfoMetadata = _createTypeLibMetadata.TypeInfoMetaData[module.GUID];
                foreach (var moduleConstant in _moduleConstants.FirstOrDefault(x => x.Key == module))
                {
                    createTypeInfoMetadata.AddConstant(moduleConstant);
                }
                SUCCEEDED(createTypeInfoMetadata.CreateTypeInfo.LayOut());
            }

            foreach (var enumType in _enums)
            {
                var createTypeInfoMetadata = _createTypeLibMetadata.TypeInfoMetaData[enumType.GUID];
                // To avoid having to define a AddConstant overload for enums, we use GetFields instead of Enum static methods
                foreach (var enumMember in enumType.GetFields(BindingFlags.Public | BindingFlags.Static).Where(f => f.IsLiteral))
                {
                    // Enum members are treated as if they are a VT_I4. In .NET an enum can be of different types but
                    // we must convert them to a VT_I4 for compliance. See:
                    // https://docs.microsoft.com/en-us/windows/win32/midl/enum
                    // https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/7b5fa59b-d8f6-4a47-9695-630d3c10363e
                    createTypeInfoMetadata.AddConstant(enumMember, VarEnum.VT_I4);
                }
                SUCCEEDED(createTypeInfoMetadata.CreateTypeInfo.LayOut());
            }

            foreach (var record in _records)
            {
                var createTypeInfoMetadata = _createTypeLibMetadata.TypeInfoMetaData[record.GUID];
                foreach (var field in record.GetFields(BindingFlags.Public | BindingFlags.Instance))
                {
                    createTypeInfoMetadata.AddVarDesc(field, VARKIND.VAR_PERINSTANCE);
                }
                SUCCEEDED(createTypeInfoMetadata.CreateTypeInfo.LayOut());
            }

            foreach(var face in _interfaces)
            {
                var createTypeInfoMetadata = _createTypeLibMetadata.TypeInfoMetaData[face.GUID];
                foreach(var method in face.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod))
                {
                    createTypeInfoMetadata.AddFuncDesc(method);
                }
                SUCCEEDED(createTypeInfoMetadata.CreateTypeInfo.LayOut());
            }
        }
    }
}
