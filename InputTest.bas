Option Explicit

Sub InputTest()
    Dim test As String
    Dim i As Integer
    
    test = InputBox("What sales value do you want to check for?", , 0)
    If test = vbNullString Then
        MsgBox ("User canceled!")
        End
    End If
    i = test

End Sub
