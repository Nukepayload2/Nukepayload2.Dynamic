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
