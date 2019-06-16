Imports System.Reflection

Public Module DynamicConversion

    Private ReadOnly _wrapperFactory As New CTypeWrapper

    ''' <summary>
    ''' Wraps the source object with the specified interface. The wrap conversion is invertible by calling 
    ''' the user-defined conversion operator. This operation can be done <see cref="CTypeDynamic(Of TargetType)(Object)"/> function. 
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

        Return CTypeWrapInternal(Of TInterface)(source, targetType, sourceType)
    End Function

    ''' <summary>
    ''' Wraps the source object with the specified interface. The wrap conversion is invertible by calling 
    ''' the user-defined conversion operator with the <see cref="CTypeDynamic(Of TargetType)(Object)"/> function. 
    ''' </summary>
    ''' <typeparam name="TSource">The type of the object to be wrapped.</typeparam>
    ''' <typeparam name="TInterface">The interface which the anonymous wrapper implements.</typeparam>
    ''' <param name="source">The object to be wrapped.</param>
    ''' <returns>If the type of <paramref name="source"/> is reference type, return the wrapped object.
    ''' Otherwise, make a copy of <paramref name="source"/> and then return the wrapped object.</returns>
    ''' <exception cref="InvalidCastException"/>
    Public Function CTypeWrap(Of TInterface As Class)(source As Object) As TInterface
        Dim targetType = GetType(TInterface)
        If Not targetType.IsInterface Then
            Throw New InvalidCastException($"You can only wrap {NameOf(source)} to an interface type.")
        End If
        If source Is Nothing Then
            Return Nothing
        End If

        Dim sourceType As Type = source.GetType

        Return CTypeWrapInternal(Of TInterface)(source, targetType, sourceType)
    End Function

    Private Function CTypeWrapInternal(Of TInterface As Class)(source As Object, targetType As Type, ByRef sourceType As Type) As TInterface
        If TypeOf source Is TInterface Then
            Return DirectCast(source, TInterface)
        End If
        Dim implementedWrapperInterface =
            GetImplementedWrapperInterface(source.GetType)

        Dim boxedSource As Object
        If implementedWrapperInterface IsNot Nothing Then
            Dim unwrappedRaw = implementedWrapperInterface.GetProperty("WrappedObject").GetValue(source)
            If unwrappedRaw Is Nothing Then
                Throw New InvalidCastException("Unwrap failed. Bad custom wrapper implementation.")
            End If
            sourceType = unwrappedRaw.GetType
            boxedSource = unwrappedRaw
        Else
            boxedSource = source
        End If

        Dim newTypeName As String = sourceType.FullName & "." & targetType.FullName
        Dim reused As Boolean = Nothing
        Dim tpBuilder = _wrapperFactory.GetCTypeWrapTypeBuilder(newTypeName, reused)
        _wrapperFactory.InitializeTypeBuilder(targetType, sourceType, reused, tpBuilder)
        Dim builtType As Type = tpBuilder.CreateTypeInfo
        Return DirectCast(Activator.CreateInstance(builtType, boxedSource), TInterface)
    End Function

    ''' <summary>
    ''' Unwraps the specified object which was wrapped with <see cref="CTypeWrap(Of TSource, TInterface)(TSource)"/> .
    ''' This function does NOT try unbox conversions, IConvertible conversions and user-defined conversion operators
    ''' when unwrapping the <paramref name="source"/> object.
    ''' But it tries those conversions
    ''' when converting the unwrapped object to <typeparamref name="TSource"/>.
    ''' </summary>
    ''' <typeparam name="TSource">The type of the object to be wrapped.</typeparam>
    ''' <param name="source">The object to be unwrapped.</param>
    ''' <returns>If the type of <paramref name="source"/> is wrapped type, return the unwrapped object.</returns>
    ''' <exception cref="InvalidCastException"/>
    Public Function CTypeUnwrap(Of TSource)(source As Object) As TSource
        If source Is Nothing Then
            Return Nothing
        End If

        Dim objType = source.GetType
        If objType.IsValueType Then
            Throw New InvalidCastException("The specified object was not wrapped by CTypeWrap.")
        End If

        Dim implementedWrapperInterface =
            GetImplementedWrapperInterface(objType)

        If implementedWrapperInterface Is Nothing Then
            Throw New InvalidCastException("The specified object was not wrapped by CTypeWrap.")
        End If

        Dim unwrappedRaw = implementedWrapperInterface.GetProperty("WrappedObject").GetValue(source)

#If SUPPORT_CTYPE_DYNAMIC Then
        Return CTypeDynamic(Of TSource)(unwrappedRaw)
#Else
        If Not GetType(TSource).IsPrimitive AndAlso TypeOf unwrappedRaw IsNot TSource Then
            Throw New PlatformNotSupportedException(CTypeWrapper.ErrorReturnValueConversionGenerationNotSupported)
        End If
        Return CType(unwrappedRaw, TSource)
#End If
    End Function

    Private Function GetImplementedWrapperInterface(objType As Type) As Type
        Return Aggregate itf In objType.GetTypeInfo.ImplementedInterfaces
               Where itf.IsGenericType AndAlso itf.IsConstructedGenericType
               Let rawGenericType = itf.GetGenericTypeDefinition
               Where rawGenericType = GetType(ICTypeWrappedObject(Of))
               Select itf Into FirstOrDefault
    End Function
End Module
