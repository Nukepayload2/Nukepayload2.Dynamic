Option Strict On

Imports System.Runtime.CompilerServices

<Assembly: InternalsVisibleTo("Nukepayload2.Dynamic.Generated")>

Module Program
    Sub Main()
        Dim testc As New TestClass With {
            .BaseValue = "B"
        }
        Dim wrapped As ITestAOrB = CTypeWrap(Of TestClass, ITestAOrB)(testc)
        wrapped.TestBase()
        Console.WriteLine(wrapped.BaseValue)
        Dim wrapped2 As ITestAAndB = CTypeWrap(Of ITestAOrB, ITestAAndB)(wrapped)
        wrapped2.TestA()
        wrapped2.TestB()
        wrapped2.CompositeValue = "D"
        Dim unwrapped = CTypeDynamic(Of TestClass)(wrapped)
        unwrapped.TestA()
        Console.WriteLine(unwrapped.CompositeValue)
    End Sub
End Module

Class TestClass
    Sub TestA()
        Console.WriteLine("Test A")
    End Sub

    Sub TestB()
        Console.WriteLine("Test B")
    End Sub

    Sub TestBase()
        Console.WriteLine("Test Base")
    End Sub

    Property BaseValue As String

    Property CompositeValue As String
End Class

Interface ITestAOrB
    Sub TestBase()
    Property BaseValue As String
End Interface

Interface ITestA
    Inherits ITestAOrB
    Sub TestA()
End Interface

Interface ITestB
    Inherits ITestAOrB
    Sub TestB()
End Interface

Interface ITestAAndB
    Inherits ITestA, ITestB
    Property CompositeValue As String
End Interface

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
