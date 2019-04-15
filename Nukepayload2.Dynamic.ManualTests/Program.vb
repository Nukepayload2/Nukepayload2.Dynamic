Option Strict On

Imports System.Runtime.CompilerServices

<Assembly: InternalsVisibleTo("Nukepayload2.Dynamic.Generated")>

Module Program
    Sub Main()
        Dim testc As New TestClass
        Dim wrapped As ITestAOrB = CTypeWrap(Of TestClass, ITestAOrB)(testc)
        wrapped.TestBase()
        Dim wrapped2 As ITestAAndB = CTypeWrap(Of ITestAOrB, ITestAAndB)(wrapped)
        wrapped2.TestA()
        wrapped2.TestB()
        Dim unwrapped = CTypeDynamic(Of TestClass)(wrapped)
        unwrapped.TestA()
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
End Class

Interface ITestAOrB
    Sub TestBase()
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
