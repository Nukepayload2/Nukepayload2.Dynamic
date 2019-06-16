# Nukepayload2.Dynamic
Provides specialized dynamic features that helps you bring existing TypeScript packages to .NET .

## Commonly used members
### CTypeWrap
- Wraps an object with the specified interface, even if the type of the object is `NotInheritable` (i.e. `sealed`).
- If an object has been wrapped, unwrap it and then wrap it with a new wrapper class.
- The wrap operation is invertible with the VB `CTypeDynamic` function or the c# `(Type)(dynamic)expression` expression.

#### Limitations
- Dynamic types are not supported yet.
- In the .NET Standard 2.x version, some dynamic type conversions may fail.

#### Usage
```vb
Dim wrapped = CTypeWrap(Of SourceType, ITargetType)(source)
```

#### Sample

##### Wrap `TestClass` with `ITest`

```vb
<Assembly: InternalsVisibleTo("Nukepayload2.Dynamic.Generated")>

Module Program
    Sub Main()
        Dim testc As New TestClass
        Dim wrapped As ITest = CTypeWrap(Of TestClass, ITest)(testc)
        wrapped.TestSub()
    End Sub
End Module

Class TestClass
    Sub TestSub()
        Console.WriteLine("Test")
    End Sub
End Class

Interface ITest
    Sub TestSub()
End Interface
```

##### Composite interface type

```vb
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
```