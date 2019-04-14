Option Strict On

Imports System.Runtime.CompilerServices
Imports Nukepayload2.Dynamic
Imports Nukepayload2.Dynamic.ManualTests

<Assembly: InternalsVisibleTo("Nukepayload2.Dynamic.Generated")>

Module Program
    Sub Main()
        Dim testc As New TestClass
        Dim converted As ITestA = CTypeWrap(Of TestClass, ITestA)(testc)
        converted.TestA()
    End Sub
End Module

Class TestClass
    Sub TestA()
        Console.WriteLine("Test A")
    End Sub

    Function aaaaa() As Long
        Dim a As Integer = 1
        Return CTypeDynamic(Of Long)(a)
    End Function

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

    Public Shared Widening Operator CType(srcObject As TestClass) As TestClassITestAOrBWrapperTemplate
        Return New TestClassITestAOrBWrapperTemplate(srcObject)
    End Operator

    Public Shared Widening Operator CType(wrapObject As TestClassITestAOrBWrapperTemplate) As TestClass
        Return wrapObject.WrappedObject
    End Operator
End Class