Option Strict On

Imports System.Runtime.CompilerServices

<Assembly: InternalsVisibleTo("Nukepayload2.Dynamic.Generated")>

Module Program
    Sub Main()
        Dim testc As New TestClass
        Dim converted As IConverted = CTypeWrap(Of TestClass, IConverted)(testc)
        converted.Test()
    End Sub
End Module

Class TestClass
    Sub Test()
        Console.WriteLine("Test")
    End Sub
End Class

Interface IConverted
    Sub Test()
End Interface
