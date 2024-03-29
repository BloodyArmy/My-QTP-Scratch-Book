'in library: TestClass.class.vbs

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

'Function GetClassInstance can be in any associated library
Public Function GetClassInstance(ByVal CLassName)
    Execute "Set GetClassInstance = New " & ClassName
End Function

'@Approach A: Using Functions to instantiate classes
'in TestClass.class.vbs
Public NewTestClass: Set NewTestClass = GetClassInstance("TestClass")

'@Approach B: Direct Instantiation
'in TestClass.class.vbs
Public NewTestClass2: Set NewTestClass2 = New TestClass