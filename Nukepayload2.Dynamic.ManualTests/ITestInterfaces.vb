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
