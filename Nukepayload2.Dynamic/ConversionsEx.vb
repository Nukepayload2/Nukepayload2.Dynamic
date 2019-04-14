Imports System.Reflection

Public Module ConversionsEx

    Private ReadOnly _wrapperFactory As New CTypeWrapper

    ''' <summary>
    ''' Wraps the source object with the specified interface.
    ''' </summary>
    ''' <typeparam name="TSource">The type of the object to be wrapped.</typeparam>
    ''' <typeparam name="TInterface">The interface which the anonymous wrapper implements.</typeparam>
    ''' <param name="source">The object to be wrapped.</param>
    ''' <returns>If the type of <paramref name="source"/> is reference type, return the wrapped object.
    ''' Otherwise, make a copy of <paramref name="source"/> and then return the wrapped object.</returns>
    ''' <exception cref="InvalidCastException"/>
    Public Function CTypeWrap(Of TSource, TInterface As Class)(source As TSource) As TInterface
        Dim targetType = GetType(TInterface)
        If Not targetType.IsInterface Then
            Throw New InvalidCastException($"You can only wrap {NameOf(source)} to an interface type.")
        End If
        If source Is Nothing Then
            Return Nothing
        End If
        Dim sourceType As Type = GetType(TSource)
        _wrapperFactory.AssertNotRefStruct(sourceType)

        Dim newTypeName As String = sourceType.FullName & "." & targetType.FullName
        Dim reused As Boolean = Nothing
        Dim tpBuilder = _wrapperFactory.GetCTypeWrapTypeBuilder(newTypeName, reused)
        _wrapperFactory.InitializeTypeBuilder(targetType, sourceType, reused, tpBuilder)
        Dim builtType As Type = tpBuilder.CreateType
        Return DirectCast(Activator.CreateInstance(builtType, source), TInterface)
    End Function

    ''' <summary>
    ''' Unwraps the specified object which was wrapped with <see cref="CTypeWrap(Of TSource, TInterface)(TSource)"/> .
    ''' </summary>
    ''' <typeparam name="TInterface">The interface which the anonymous wrapper implements.</typeparam>
    ''' <typeparam name="TSource">The type of the object to be wrapped.</typeparam>
    ''' <param name="source">The object to be unwrapped.</param>
    ''' <returns>If the type of <paramref name="source"/> is wrapped type, return the unwrapped object.</returns>
    ''' <exception cref="InvalidCastException"/>
    Public Function CTypeUnwrap(Of TInterface As Class, TSource)(source As TInterface) As TSource
        If source Is Nothing Then
            Return Nothing
        End If

        Dim objType = source.GetType
        If objType.IsValueType Then
            Throw New InvalidCastException("The specified object was not wrapped by CTypeWrap.")
        End If

        Dim implementedWrappedObject =
            Aggregate itf In objType.GetTypeInfo.ImplementedInterfaces
            Where itf.IsGenericType AndAlso itf.IsConstructedGenericType
            Let rawGenericType = itf.GetGenericTypeDefinition
            Where rawGenericType = GetType(ICTypeWrappedObject(Of)) Into FirstOrDefault

        If implementedWrappedObject Is Nothing Then
            Throw New InvalidCastException("The specified object was not wrapped by CTypeWrap.")
        End If

        Dim wrapperInterface = implementedWrappedObject.itf
        Dim unwrappedRaw = wrapperInterface.GetProperty("WrappedObject").GetValue(source)

        Return CTypeDynamic(Of TSource)(unwrappedRaw)
    End Function

End Module
