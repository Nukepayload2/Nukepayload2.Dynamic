Option Strict On

Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.CompilerServices

Public Module ConversionsEx

    Private Const CTypeWrapAssemblyName = "Nukepayload2.Dynamic.Generated"
    Private Const RefStructAttrName = "System.Runtime.CompilerServices.IsByRefLikeAttribute"
    Private _cTypeWrapAssemblyBuilder As AssemblyBuilder
    Private _cTypeWrapModuleBuilder As ModuleBuilder
    Private _cTypeWrapTypeBuilders As Dictionary(Of String, TypeBuilder)

    Private ReadOnly Property CTypeWrapAssemblyBuilder As AssemblyBuilder
        <MethodImpl(MethodImplOptions.Synchronized)>
        Get
            If _cTypeWrapAssemblyBuilder Is Nothing Then
                _cTypeWrapAssemblyBuilder =
                    AppDomain.CurrentDomain.DefineDynamicAssembly(
                        New AssemblyName(CTypeWrapAssemblyName),
                        AssemblyBuilderAccess.Run)
            End If
            Return _cTypeWrapAssemblyBuilder
        End Get
    End Property

    Private ReadOnly Property CTypeWrapModuleBuilder As ModuleBuilder
        <MethodImpl(MethodImplOptions.Synchronized)>
        Get
            If _cTypeWrapModuleBuilder Is Nothing Then
                _cTypeWrapModuleBuilder = CTypeWrapAssemblyBuilder.DefineDynamicModule(CTypeWrapAssemblyName)
            End If
            Return _cTypeWrapModuleBuilder
        End Get
    End Property

    <MethodImpl(MethodImplOptions.Synchronized)>
    Private Function GetCTypeWrapTypeBuilder(key As String, Optional ByRef isReused As Boolean = False) As TypeBuilder
        If _cTypeWrapTypeBuilders Is Nothing Then
            _cTypeWrapTypeBuilders = New Dictionary(Of String, TypeBuilder)
        End If
        Dim tpBuilder As TypeBuilder = Nothing
        If _cTypeWrapTypeBuilders.TryGetValue(key, tpBuilder) Then
            isReused = True
        Else
            tpBuilder = CTypeWrapModuleBuilder.DefineType(
                key, TypeAttributes.Public Or TypeAttributes.Sealed)
            _cTypeWrapTypeBuilders.Add(key, tpBuilder)
        End If
        Return tpBuilder
    End Function

    Function CTypeWrap(Of TSource, TInterface As Class)(source As TSource) As TInterface
        Dim targetType = GetType(TInterface)
        If Not targetType.IsInterface Then
            Throw New InvalidCastException($"You can only wrap {NameOf(source)} to an interface type.")
        End If
        If source Is Nothing Then
            Return Nothing
        End If
        Dim sourceType As Type = GetType(TSource)
        Dim isByRefLikeAttribute = From attr In sourceType.GetCustomAttributes
                                   Where attr.GetType.FullName = RefStructAttrName
        If sourceType.IsValueType AndAlso isByRefLikeAttribute.Any Then
            Throw New InvalidCastException($"You cann't box a ref struct.")
        End If
        Dim newTypeName As String = sourceType.FullName & "." & targetType.FullName
        Dim reused As Boolean = Nothing
        Dim tpBuilder = GetCTypeWrapTypeBuilder(newTypeName, reused)
        If Not reused Then
            tpBuilder.AddInterfaceImplementation(targetType)
            Dim backingFld = tpBuilder.DefineField("_wrapBackingField", sourceType, FieldAttributes.Private)
            GenerateMethodWrappers(targetType, sourceType, tpBuilder, backingFld)
            GeneratePropertyWrappers(targetType, tpBuilder)
            GenerateEventWrappers(targetType, tpBuilder)
            GenerateConstructorWrapper(sourceType, tpBuilder, backingFld)
        End If
        Dim builtType As Type = tpBuilder.CreateType
        Return DirectCast(Activator.CreateInstance(builtType, source), TInterface)
    End Function

    Private Sub GenerateConstructorWrapper(sourceType As Type, tpBuilder As TypeBuilder, backingFld As FieldBuilder)
        Dim ctorBuilder = tpBuilder.DefineConstructor(MethodAttributes.Public, CallingConventions.Standard, {sourceType})
        Dim ctorBody = ctorBuilder.GetILGenerator
        ctorBody.Emit(OpCodes.Ldarg_0)
        ctorBody.Emit(OpCodes.Ldarg_1)
        ctorBody.Emit(OpCodes.Stfld, backingFld)
        ctorBody.Emit(OpCodes.Ret)
    End Sub

    Private Sub GenerateEventWrappers(targetType As Type, tpBuilder As TypeBuilder)
        For Each evt In targetType.GetEvents
            tpBuilder.DefineEvent(evt.Name, EventAttributes.None, evt.EventHandlerType)
        Next
    End Sub

    Private Sub GeneratePropertyWrappers(targetType As Type, tpBuilder As TypeBuilder)
        For Each interfaceProperty In targetType.GetProperties
            tpBuilder.DefineProperty(interfaceProperty.Name, PropertyAttributes.None,
                                     interfaceProperty.PropertyType,
                                     Aggregate p In interfaceProperty.GetIndexParameters
                                     Select p.ParameterType Into ToArray)
        Next
    End Sub

    Private Sub GenerateMethodWrappers(targetType As Type, sourceType As Type, tpBuilder As TypeBuilder, backingFld As FieldBuilder)
        For Each interfaceMethod In targetType.GetMethods
            Dim interfaceMethodParams = interfaceMethod.GetParameters
            Dim interfaceMethodParamTypes = Aggregate p In interfaceMethodParams Select p.ParameterType Into ToArray
            Dim wrappedMethod As MethodInfo = sourceType.GetMethod(interfaceMethod.Name, interfaceMethodParamTypes)
            If wrappedMethod Is Nothing Then
                Throw New InvalidCastException($"Unable to cast {sourceType.FullName} to {targetType.FullName}")
            End If
            Dim methodAttr = MethodAttributes.Public Or MethodAttributes.Virtual Or
                MethodAttributes.CheckAccessOnOverride Or MethodAttributes.Final Or MethodAttributes.NewSlot
            If interfaceMethod.IsSpecialName Then
                ' Get, Set, Add, Remove and Raise
                methodAttr = methodAttr Or MethodAttributes.SpecialName
            End If
            Dim wrapperMethodBuilder = tpBuilder.DefineMethod(
                targetType.Name & "_" & interfaceMethod.Name,
                methodAttr,
                CallingConventions.HasThis,
                interfaceMethod.ReturnType,
                interfaceMethodParamTypes)
            Dim bodyGen = wrapperMethodBuilder.GetILGenerator
            bodyGen.Emit(OpCodes.Ldarg_0)
            bodyGen.Emit(OpCodes.Ldfld, backingFld)
            For i = 1 To interfaceMethodParams.Length
                bodyGen.Emit(OpCodes.Ldarg, i)
            Next
            bodyGen.EmitCall(OpCodes.Callvirt, wrappedMethod, interfaceMethodParamTypes)
            bodyGen.Emit(OpCodes.Ret)
            tpBuilder.DefineMethodOverride(wrapperMethodBuilder, interfaceMethod)
        Next
    End Sub
End Module
