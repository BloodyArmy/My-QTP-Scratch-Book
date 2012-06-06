'in Action: TestClass

Class TestClass

   Public Sub Method()
	Print "In Method()"
   End Sub

   Public Sub Class_Initialize()
	Print "In Class_Initialize()"
   End Sub

   Public Sub Class_Terminate()
	Print "In Class_Terminate()"
   End Sub

End Class



Public Function GetClassInstance(ByVal ClassName)
    Execute "Set GetClassInstance = New " & ClassName
End Function
 
'call made in ClassA
Set NewTestClass = GetClassInstance("TestClass")
NewTestClass.Method()
Set NewTestClass = Nothing
 
'call made in ClassB
Set NewTestClass = GetClassInstance("TestClass")
NewTestClass.Method()
Set NewTestClass = Nothing
 
'call made in Action1
Set NewTestClass = GetClassInstance("TestClass")
NewTestClass.Method()
Set NewTestClass = Nothing