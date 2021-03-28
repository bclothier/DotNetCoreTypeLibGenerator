using DotNetCoreTypeLibGenerator.Abstract;
using DotNetCoreTypeLibGenerator.Extensions;
using DotNetCoreTypeLibGenerator.Helpers;
using DotNetCoreTypeLibGenerator.Unmanaged;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using static Vanara.PInvoke.OleAut32;
using static DotNetCoreTypeLibGenerator.Helpers.HResultHelper;

using CALLCONV = System.Runtime.InteropServices.ComTypes.CALLCONV;

namespace DotNetCoreTypeLibGenerator
{
    internal struct CreateTypeInfoAttributes
    {
        public string Name { get; set; }
        public TYPEFLAGS Flags { get; set; }
        public ushort MajorVersion { get; set; }
        public ushort MinorVersion { get; set; }
        public string DocString { get; set; }
        public uint HelpContext { get; set; }
        public uint HelpStringContext { get; set; }
    }

    internal class CreateTypeInfoMetadata
    {
        private readonly ITypeLibGenerationEventSink _sink;
        private readonly IReferencedTypeLibrariesProvider _referencedTypeLibProvider;
        private MemberIdDispenser _memIdDispenser = new MemberIdDispenser();

        public TYPEKIND TypeKind { get; }
        public Guid Guid { get; }
        public Type Type { get; }
        public ICreateTypeInfo2 CreateTypeInfo { get; }
        public ITypeInfo2 TypeInfo => (ITypeInfo2)CreateTypeInfo;
        private Dictionary<int, CreateTypeInfoMemberMetadata> _memberMetadata;
        public IReadOnlyDictionary<int, CreateTypeInfoMemberMetadata> MemberMetaData => _memberMetadata;
        public CreateTypeInfoAttributes Attributes { get; }

        public CreateTypeInfoMetadata(TYPEKIND typeKind, Type type, ICreateTypeInfo2 createTypeInfo, ITypeLibGenerationEventSink sink, IReferencedTypeLibrariesProvider referencedTypeLibProvider)
        {
            _sink = sink;

            TypeKind = typeKind;
            Type = type;
            CreateTypeInfo = createTypeInfo;
            Guid = AttributeHelper.GetGuid(type);

            if(Guid == null || Guid == Guid.Empty)
            {
                throw new ArgumentException($"The type {type.FullName} does not have a GUID assigned. The type must have a Guid attribute.");
            }

            createTypeInfo.SetGuid(Guid);
            createTypeInfo.SetName(type.Name);

            Attributes = new CreateTypeInfoAttributes
            {
                Name = AttributeHelper.GetComAliasName(Type) ?? Type.Name
            };

            _memberMetadata = new Dictionary<int, CreateTypeInfoMemberMetadata>();
            _referencedTypeLibProvider = referencedTypeLibProvider;
        }

        public void AddFuncDesc(MethodInfo methodInfo)
        {
            // TODO: Handle both presence / absence of the PreserveSig attribute. Currently, the code
            // TODO: acts as if the attribute was present but in reality, we should default to transforming
            // TODO: the signature returning a HRESULT and taking the return parameter as the final parameter.
            var invKind = InferInvKind(methodInfo);
            var funcName = methodInfo.Name;
            if(invKind!= INVOKEKIND.INVOKE_FUNC)
            {
                var propertyInfo = methodInfo.DeclaringType.GetProperty(methodInfo.Name.Substring("xxx_".Length));
                if (propertyInfo != null)
                {
                    funcName = propertyInfo.Name;
                }
            }
            var member = AddMember(funcName, methodInfo, methodInfo.ReturnType.GetVarEnum());
            var names = new List<string>
            {
                funcName
            };

            var parameters = methodInfo.GetParameters();

            var reqParameters = parameters.Where(x => !x.IsOptional);
            var optParameters = parameters.Where(x => x.IsOptional);
            var retValParameter = parameters.FirstOrDefault(x => x.IsRetval) ?? methodInfo.ReturnParameter;

            var cParams = reqParameters.Count();
            var cParamsOpt = optParameters.Count();
            var elementSize = Marshal.SizeOf<ELEMDESC>();
            
            // We want to free all unmanaged memory allocations at the end of the procedure because
            // there could be multiple memory allocations for parameters and we cannot deallocate 
            // until we are done with all the parameters description.
            using var handles = new DisposableList<SafeHandle>();

            UnmanagedMemory.UsingCoTaskMem(elementSize * (cParams + cParamsOpt), ptr => {
                void ProcessParameter(ParameterInfo p)
                {
                    var iteratorPtr = ptr;

                    // TODO: Handle indexed parameters -- C# does not support it but VB.NET does and therefore
                    // TODO: we would need the parameters' name added but not the retVal which must be anonymous
                    if (invKind != INVOKEKIND.INVOKE_PROPERTYPUT && invKind != INVOKEKIND.INVOKE_PROPERTYPUTREF)
                    {
                        names.Add(p.Name);
                    }
                    var elemDescParam = new ELEMDESC();
                    var paramFlags = CalculateParameterFlags(p, out var defaultValue);

                    elemDescParam.tdesc.vt = (short)p.ParameterType.GetVarEnum();
                    elemDescParam.desc.paramdesc.wParamFlags = paramFlags;
                    if (defaultValue != null)
                    {
                        var defaultHandle = new PARAMDESCEX(defaultValue).GetHandle();
                        handles.Add(defaultHandle);
                        elemDescParam.desc.paramdesc.lpVarValue = defaultHandle.DangerousGetHandle();
                    }
                    else
                    {
                        elemDescParam.desc.paramdesc.lpVarValue = IntPtr.Zero;
                    }
                    Marshal.StructureToPtr(elemDescParam, iteratorPtr, false);
                    iteratorPtr += elementSize;
                };

                foreach (var p in reqParameters)
                {
                    ProcessParameter(p);
                }

                foreach(var p in optParameters)
                {
                    ProcessParameter(p);
                }

                var elemDescReturn = new ELEMDESC();
                if (retValParameter != null)
                {
                    elemDescReturn.tdesc.vt = (short)retValParameter.ParameterType.GetVarEnum();
                }

                // See:
                // https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/d3349d25-e11d-4095-ba86-de3fda178c4e
                var funcDesc = new FUNCDESC
                {
                    memid = member.MemberId,
                    cScodes = 0,
                    lprgelemdescParam = ptr,
                    funckind = FUNCKIND.FUNC_PUREVIRTUAL,
                    invkind = invKind,
                    // Automation assumes STDCALL calling conventions so that should be the only convention we'll use
                    callconv = CALLCONV.CC_STDCALL,
                    cParams = (short)cParams,
                    cParamsOpt = (short)cParamsOpt,
                    oVft = 0,
                    elemdescFunc = elemDescReturn,
                    wFuncFlags = 0
                };

                SUCCEEDED(CreateTypeInfo.AddFuncDesc((uint)member.MemberId, funcDesc));
                SUCCEEDED(CreateTypeInfo.SetFuncAndParamNames((uint)member.MemberId, names.ToArray(), (uint)names.Count));
            });
        }

        /// <summary>
        /// There is no nice, tidy way to check if a <see cref="MethodInfo"/> is a property accessor or a true method.
        /// Therefore, we have to do extra work to check if it's a property accessor, and assume it's a regular method
        /// as the default. 
        /// </summary>
        /// <param name="methodInfo">The method to determine whether it's a method or a proeprty accessor.</param>
        /// <returns>The corresponding <see cref="INVOKEKIND"/> value</returns>
        private INVOKEKIND InferInvKind(MethodInfo methodInfo)
        {
            if ((methodInfo.Attributes & MethodAttributes.SpecialName) == MethodAttributes.SpecialName)
            {
                if (methodInfo.Name.StartsWith("get_"))
                {
                    return INVOKEKIND.INVOKE_PROPERTYGET;
                }
                else if (methodInfo.Name.StartsWith("set_"))
                {
                    if (methodInfo.ReturnType.IsValueType)
                    {
                        return INVOKEKIND.INVOKE_PROPERTYPUT;
                    }
                    else
                    {
                        return INVOKEKIND.INVOKE_PROPERTYPUTREF;
                    }
                }
            }
            return INVOKEKIND.INVOKE_FUNC;
        }

        /// <summary>
        /// Set the <see cref="PARAMFLAG"/> based on the <see cref="ParameterInfo"/> contents and 
        /// presence/absence of the various parameter attributes such as <see cref="InAttribute"/>
        /// or <see cref="OutAttribute"/>. 
        /// </summary>
        /// <param name="param">The <see cref="ParameterInfo"/> to extract the flags.</param>
        /// <param name="defaultValue">If the <paramref name="param"/> has a default value, it is returned.</param>
        /// <returns>The corresponding <see cref="PARAMFLAG"/> flags.</returns>
        private PARAMFLAG CalculateParameterFlags(ParameterInfo param, out object defaultValue)
        {
            // TODO: Need to write tests for this procedure
            // TODO: We need to be more smarter than just checking the proeprties & attributes as the source code
            // TODO: may not have any attributes decorated and the properties may not be accurate, so we need to 
            // TODO: derive the In/Out/etc. based on how it's passed in/out if we do not have any explicit information
            // TODO: from the attributes or the properties. 
            // TODO: However, how do we know the property is wrong? This also complicates the override logic. 
            var paramFlags = PARAMFLAG.PARAMFLAG_NONE;

            if (AttributeHelper.IsInParameter(param) || param.IsIn)
            {
                paramFlags |= PARAMFLAG.PARAMFLAG_FIN;
            }
            if (AttributeHelper.IsOutParameter(param) || param.IsOut)
            {
                paramFlags |= PARAMFLAG.PARAMFLAG_FOUT;
            }
            if (AttributeHelper.IsOptionalParameter(param) || param.IsOptional)
            {
                paramFlags |= PARAMFLAG.PARAMFLAG_FOPT;
            }
            
            defaultValue = AttributeHelper.HasDefaultValue(param) ?? param.DefaultValue;
            if (defaultValue == DBNull.Value)
            {
                defaultValue = null;
            }

            if (defaultValue != null)
            {
                paramFlags |= PARAMFLAG.PARAMFLAG_FHASDEFAULT;
            }

            if(param.IsRetval)
            {
                paramFlags |= PARAMFLAG.PARAMFLAG_FRETVAL;
            }

            return paramFlags;
        }

        public void AddVarDesc(FieldInfo fieldInfo, VARKIND varKind)
        {
            AddVarDesc(fieldInfo, varKind, fieldInfo.FieldType.GetVarEnum());
        }

        public void AddVarDesc(FieldInfo fieldInfo, VARKIND varKind, VarEnum varEnum)
        {
            var member = AddMember(fieldInfo.Name, fieldInfo, varEnum);
            var varDesc = new VARDESC
            {
                varkind = varKind,
                memid = member.MemberId
            };
            varDesc.elemdescVar.tdesc.vt = (short)varEnum;
            SUCCEEDED(CreateTypeInfo.AddVarDesc((uint)member.MemberId, varDesc));
            SUCCEEDED(CreateTypeInfo.SetVarName((uint)member.MemberId, fieldInfo.Name));
        }

        private CreateTypeInfoMemberMetadata AddMember(string specifiedName, MemberInfo memberInfo, VarEnum varEnum)
        {
            var assignedMemId = _memIdDispenser.GetMemberId(specifiedName ?? memberInfo.Name);
            var memberMetadata = new CreateTypeInfoMemberMetadata(assignedMemId, memberInfo, varEnum);
            if (!_memberMetadata.ContainsKey(assignedMemId))
            {
                _memberMetadata.Add(assignedMemId, memberMetadata);
            }
            return memberMetadata;
        }

        public void AddConstant(FieldInfo constant)
        {
            AddConstant(constant, constant.FieldType.GetVarEnum());
        }

        public void AddConstant(FieldInfo constant, VarEnum varEnum)
        {
            var constantName = AttributeHelper.GetComAliasName(constant) ?? constant.Name;
            var constantValue = (constant.IsLiteral && !constant.IsInitOnly) ? constant.GetRawConstantValue() : constant.GetValue(null);
            var assignedMemId = _memIdDispenser.GetMemberId(constantName);
            var desc = new VARDESC()
            {
                memid = assignedMemId,
                varkind = VARKIND.VAR_CONST,
            };
            UnmanagedMemory.UsingVariant(constantValue, ptr =>
            {
                desc.desc.lpvarValue = ptr;
                desc.elemdescVar.tdesc.vt = (short)varEnum;
                SUCCEEDED(CreateTypeInfo.AddVarDesc((uint)assignedMemId, desc));
                SUCCEEDED(CreateTypeInfo.SetVarName((uint)assignedMemId, constantName));
            });
        }
    }
}
