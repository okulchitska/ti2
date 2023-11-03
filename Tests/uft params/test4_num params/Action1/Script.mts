Option Explicit

Dim MyMsgBox
Set MyMsgBox = DotNetFactory.CreateInstance("System.Windows.Forms.MessageBox", "System.Windows.Forms")

Dim a, b, result

a = Parameter("aA")
b = Parameter("aB")
result = a + b

Msgbox "Result: " & a & " + "& b & " = " & result
