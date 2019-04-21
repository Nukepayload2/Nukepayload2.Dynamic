Imports System.Runtime.CompilerServices

<Discardable>
Class TestClassITestAOrBWrapperTemplate
    Implements ICTypeWrappedObject(Of TestClass), ITestAOrB

    Private _backingField As TestClass

    Sub New(backingField As TestClass)
        _backingField = backingField
    End Sub

    Public ReadOnly Property WrappedObject As TestClass Implements ICTypeWrappedObject(Of TestClass).WrappedObject
        Get
            Return _backingField
        End Get
    End Property

    Public Property BaseValue As String Implements ITestAOrB.BaseValue
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As String)
            Throw New NotImplementedException()
        End Set
    End Property

    Public Sub TestBase() Implements ITestAOrB.TestBase
        _backingField.TestBase()
    End Sub

    ' .method public specialname static class Nukepayload2.Dynamic.ManualTests.TestClassITestAOrBWrapperTemplate op_Implicit (
    '     class Nukepayload2.Dynamic.ManualTests.TestClass srcObject
    ' ) cil managed 
    Public Shared Widening Operator CType(srcObject As TestClass) As TestClassITestAOrBWrapperTemplate
        ' ldarg0
        ' newobj TestClassITestAOrBWrapperTemplate
        ' ret
        Return New TestClassITestAOrBWrapperTemplate(srcObject)
    End Operator

    Public Shared Widening Operator CType(wrapObject As TestClassITestAOrBWrapperTemplate) As TestClass
        ' ldarg0
        ' call WrappedObject
        ' ret
        Return wrapObject.WrappedObject
    End Operator
End Class
