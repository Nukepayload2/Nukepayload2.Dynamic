Option Strict On

Imports System.Runtime.CompilerServices

<Assembly: InternalsVisibleTo("Nukepayload2.Dynamic.Generated")>

Module Program
    Sub Main()
        Dim testc As New TestClass With {
            .BaseValue = "B"
        }
        Console.WriteLine("testc.BaseValue = ""B""")
        Dim wrapped As ITestAOrB = CTypeWrap(Of TestClass, ITestAOrB)(testc)
        wrapped.TestBase()
        Console.WriteLine("Get testc.BaseValue")
        Console.WriteLine(wrapped.BaseValue)
        Dim wrapped2 As ITestAAndB = CTypeWrap(Of ITestAOrB, ITestAAndB)(wrapped)
        wrapped2.TestA()
        wrapped2.TestB()
        Console.WriteLine("wrapped2.CompositeValue = ""D""")
        wrapped2.CompositeValue = "D"
        Dim unwrapped = CTypeDynamic(Of TestClass)(wrapped)
        unwrapped.TestA()
        Console.WriteLine("Get wrapped2.CompositeValue")
        Console.WriteLine(unwrapped.CompositeValue)
    End Sub
End Module
