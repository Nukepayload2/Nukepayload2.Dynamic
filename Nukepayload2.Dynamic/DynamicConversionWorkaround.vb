Imports System.ComponentModel
Imports System.Dynamic
Imports System.Reflection
Imports System.Runtime.CompilerServices

Public Module Conversion

    Friend Const ErrorReturnValueConversionGenerationNotSupported = "The current runtime doesn't support generating the conversion between dynamic return values. Consider using the .NET Framework 4.6.1 assembly of this package."

    ''' <summary>
    ''' The .NET Standard 2.0 version of VB runtime doesn't have CTypeDynamic.
    ''' This is a less powerful replacement of CTypeDynamic.
    ''' </summary>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Function CTypeDynamic(Of T)(obj As Object) As T
        If obj Is Nothing Then
            Return Nothing
        End If
        Dim srcType = obj.GetType
        Dim destType = GetType(T)

        ' Same type
        If destType.IsAssignableFrom(srcType) Then
            Return DirectCast(obj, T)
        End If

        ' IConvertible conversion
        If TypeOf srcType Is IConvertible AndAlso GetType(IConvertible).IsAssignableFrom(destType) Then
            Return DirectCast(Convert.ChangeType(obj, destType), T)
        End If

        ' DLR conversion is not supported.
        ' Throw exception for this case.
        If TypeOf obj Is IDynamicMetaObjectProvider OrElse GetType(IDynamicMetaObjectProvider).IsAssignableFrom(destType) Then
            Throw New PlatformNotSupportedException(ErrorReturnValueConversionGenerationNotSupported)
        End If

        ' User defined conversion
        Dim convOperators = GetUserDefinedConversionOperator(srcType, destType)
        If convOperators IsNot Nothing Then
            Return DirectCast(convOperators.Invoke(Nothing, {obj}), T)
        End If

        Throw New InvalidCastException($"Unable to cast {obj.GetType.ToString} to {GetType(T).ToString}.")
    End Function

    Private Function GetUserDefinedConversionOperator(srcType As Type, destType As Type) As MethodInfo
        ' This operation is heavy. We need a cache here.
        Return Aggregate op In srcType.GetConversionOperators.Concat(destType.GetConversionOperators)
               Where op.GetParameters(0).ParameterType = srcType AndAlso op.ReturnType = destType
               Into FirstOrDefault
    End Function

    <Extension>
    Private Function GetConversionOperators(srcType As Type) As IEnumerable(Of MethodInfo)
        Const WideningOperatorName = "op_Implicit"
        Const NarrowingOperatorName = "op_Explicit"

        Return From m In srcType.GetMethods
               Where m.IsSpecialName AndAlso m.IsStatic AndAlso
                     m.Name = WideningOperatorName OrElse m.Name = NarrowingOperatorName
    End Function
End Module
