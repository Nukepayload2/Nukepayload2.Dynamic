Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.CompilerServices

Friend Class CTypeWrapper

#If Not SUPPORT_CTYPE_DYNAMIC Then
    Friend Const ErrorReturnValueConversionGenerationNotSupported = "The current runtime doesn't support generating the dynamic conversion between return values. Consider using the .NET Framework 4.6.1 assembly of this package."
#End If

    Private Const CTypeWrapAssemblyName = "Nukepayload2.Dynamic.Generated"
    Private Const RefStructAttrName = "System.Runtime.CompilerServices.IsByRefLikeAttribute"
    Private _cTypeWrapAssemblyBuilder As AssemblyBuilder
    Private _cTypeWrapModuleBuilder As ModuleBuilder
    Private _cTypeWrapTypeBuilders As Dictionary(Of String, TypeBuilder)

    Public ReadOnly Property CTypeWrapAssemblyBuilder As AssemblyBuilder
        <MethodImpl(MethodImplOptions.Synchronized)>
        Get
            If _cTypeWrapAssemblyBuilder Is Nothing Then
                _cTypeWrapAssemblyBuilder =
                    AssemblyBuilder.DefineDynamicAssembly(
                        New AssemblyName(CTypeWrapAssemblyName),
                        AssemblyBuilderAccess.Run)
            End If
            Return _cTypeWrapAssemblyBuilder
        End Get
    End Property

    Public ReadOnly Property CTypeWrapModuleBuilder As ModuleBuilder
        <MethodImpl(MethodImplOptions.Synchronized)>
        Get
            If _cTypeWrapModuleBuilder Is Nothing Then
                _cTypeWrapModuleBuilder = CTypeWrapAssemblyBuilder.DefineDynamicModule(CTypeWrapAssemblyName)
            End If
            Return _cTypeWrapModuleBuilder
        End Get
    End Property

    <MethodImpl(MethodImplOptions.Synchronized)>
    Public Function GetCTypeWrapTypeBuilder(key As String, Optional ByRef isReused As Boolean = False) As TypeBuilder
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

    Public Sub AssertNotRefStruct(sourceType As Type)
        Dim isByRefLikeAttribute = From attr In sourceType.GetCustomAttributes
                                   Where attr.GetType.FullName = RefStructAttrName
        If sourceType.IsValueType AndAlso isByRefLikeAttribute.Any Then
            Throw New InvalidCastException($"You cann't box a ref struct.")
        End If
    End Sub

    Public Sub InitializeTypeBuilder(targetType As Type, sourceType As Type, reused As Boolean, tpBuilder As TypeBuilder)
        If Not reused Then
            tpBuilder.AddInterfaceImplementation(targetType)
            Dim wrapperInterface = GetType(ICTypeWrappedObject(Of)).MakeGenericType(sourceType)
            tpBuilder.AddInterfaceImplementation(wrapperInterface)
            Dim backingFld = tpBuilder.DefineField("_wrapBackingField", sourceType, FieldAttributes.Private)
            GenerateMethodWrappers(targetType, sourceType, tpBuilder, backingFld)
            GenerateWrapperBackingFieldGetMethod(wrapperInterface, tpBuilder, backingFld)
            GeneratePropertyWrappers(targetType, tpBuilder)
            GeneratePropertyWrappers(wrapperInterface, tpBuilder)
            GenerateEventWrappers(targetType, tpBuilder)
            GenerateConstructorWrapper(sourceType, tpBuilder, backingFld)
            GenerateWideningCTypeOperator(sourceType, wrapperInterface, tpBuilder)
        End If
    End Sub

    Private Sub GenerateWideningCTypeOperator(sourceType As Type, wrapperInterface As Type, tpBuilder As TypeBuilder)
        Dim methodAttr = MethodAttributes.Public Or MethodAttributes.SpecialName Or
            MethodAttributes.Static
        Dim cTypeUnwrapOperator = tpBuilder.DefineMethod("op_Implicit", methodAttr, sourceType, {tpBuilder})
        Dim il = cTypeUnwrapOperator.GetILGenerator
        il.Emit(OpCodes.Ldarg_0)
        il.Emit(OpCodes.Callvirt, wrapperInterface.GetProperty(NameOf(ICTypeWrappedObject(Of Object).WrappedObject)).GetMethod)
        il.Emit(OpCodes.Ret)
    End Sub

    Private Sub GenerateConstructorWrapper(sourceType As Type, tpBuilder As TypeBuilder, backingFld As FieldBuilder)
        Dim ctorBuilder = tpBuilder.DefineConstructor(MethodAttributes.Public, CallingConventions.Standard, {sourceType})
        Dim ctorBody = ctorBuilder.GetILGenerator
        ctorBody.Emit(OpCodes.Ldarg_0)
        ctorBody.Emit(OpCodes.Ldarg_1)
        ctorBody.Emit(OpCodes.Stfld, backingFld)
        ctorBody.Emit(OpCodes.Ret)
    End Sub

    Private Sub GenerateEventWrappers(targetType As Type, tpBuilder As TypeBuilder, Optional processedTypes As HashSet(Of Type) = Nothing)
        If processedTypes Is Nothing Then
            processedTypes = New HashSet(Of Type)
        ElseIf processedTypes.Contains(targetType) Then
            Return
        End If
        For Each evt In targetType.GetEvents
            tpBuilder.DefineEvent(evt.Name, EventAttributes.None, evt.EventHandlerType)
        Next
        processedTypes.Add(targetType)
        For Each baseInterface In targetType.GetInterfaces
            GenerateEventWrappers(baseInterface, tpBuilder, processedTypes)
        Next
    End Sub

    Private Sub GeneratePropertyWrappers(targetType As Type, tpBuilder As TypeBuilder, Optional processedTypes As HashSet(Of Type) = Nothing)
        If processedTypes Is Nothing Then
            processedTypes = New HashSet(Of Type)
        ElseIf processedTypes.Contains(targetType) Then
            Return
        End If
        For Each interfaceProperty In targetType.GetProperties
            tpBuilder.DefineProperty(targetType.Name & "_" & interfaceProperty.Name, PropertyAttributes.None,
                                     interfaceProperty.PropertyType,
                                     Aggregate p In interfaceProperty.GetIndexParameters
                                     Select p.ParameterType Into ToArray)
        Next
        processedTypes.Add(targetType)
        For Each baseInterface In targetType.GetInterfaces
            GeneratePropertyWrappers(baseInterface, tpBuilder, processedTypes)
        Next
    End Sub

    Private Sub GenerateMethodWrappers(targetType As Type, sourceType As Type, tpBuilder As TypeBuilder, backingFld As FieldBuilder, Optional processedTypes As HashSet(Of Type) = Nothing)
        If processedTypes Is Nothing Then
            processedTypes = New HashSet(Of Type)
        ElseIf processedTypes.Contains(targetType) Then
            Return
        End If
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
            GenerateWrapperMethod(targetType, tpBuilder, backingFld, interfaceMethod, interfaceMethodParams,
                                  interfaceMethodParamTypes, wrappedMethod, methodAttr)
        Next
        processedTypes.Add(targetType)
        For Each baseInterface In targetType.GetInterfaces
            GenerateMethodWrappers(baseInterface, sourceType, tpBuilder, backingFld, processedTypes)
        Next
    End Sub

    Private Sub GenerateWrapperMethod(targetType As Type, tpBuilder As TypeBuilder,
                                      backingFld As FieldBuilder, interfaceMethod As MethodInfo,
                                      interfaceMethodParams() As ParameterInfo, interfaceMethodParamTypes() As Type,
                                      wrappedMethod As MethodInfo, methodAttr As MethodAttributes)
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
        bodyGen.Emit(OpCodes.Callvirt, wrappedMethod)
        If interfaceMethod.ReturnType <> wrappedMethod.ReturnType Then
            If wrappedMethod.ReturnType = GetType(Void) Then
                If wrappedMethod.ReturnType.IsValueType Then
                    bodyGen.Emit(OpCodes.Newobj, wrappedMethod.ReturnType)
                Else
                    bodyGen.Emit(OpCodes.Ldnull)
                End If
            ElseIf interfaceMethod.ReturnType = GetType(Void) Then
                bodyGen.Emit(OpCodes.Pop)
            Else
#If SUPPORT_CTYPE_DYNAMIC Then
                Dim vbCTypeDynamic = GetType(Conversion).GetMethod("CTypeDynamic", {GetType(Object)}).
                    MakeGenericMethod(interfaceMethod.ReturnType)
                bodyGen.Emit(OpCodes.Call, vbCTypeDynamic)
#Else
                Throw New PlatformNotSupportedException(ErrorReturnValueConversionGenerationNotSupported)
#End If
            End If
        End If
        bodyGen.Emit(OpCodes.Ret)
        tpBuilder.DefineMethodOverride(wrapperMethodBuilder, interfaceMethod)
    End Sub

    Private Sub GenerateWrapperBackingFieldGetMethod(targetType As Type, tpBuilder As TypeBuilder,
                                      backingFld As FieldBuilder)
        Dim interfaceMethod As MethodInfo =
            targetType.GetProperty(NameOf(ICTypeWrappedObject(Of Object).WrappedObject)).GetMethod
        Dim methodAttr = MethodAttributes.Public Or MethodAttributes.Virtual Or
            MethodAttributes.CheckAccessOnOverride Or MethodAttributes.Final Or MethodAttributes.NewSlot
        Dim wrapperMethodBuilder = tpBuilder.DefineMethod(
            targetType.Name & "_" & interfaceMethod.Name,
            methodAttr,
            CallingConventions.HasThis,
            interfaceMethod.ReturnType,
            Type.EmptyTypes)
        Dim bodyGen = wrapperMethodBuilder.GetILGenerator
        bodyGen.Emit(OpCodes.Ldarg_0)
        bodyGen.Emit(OpCodes.Ldfld, backingFld)
        bodyGen.Emit(OpCodes.Ret)
        tpBuilder.DefineMethodOverride(wrapperMethodBuilder, interfaceMethod)
    End Sub

End Class
