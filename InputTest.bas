Option Explicit
'Function to check user input type/cancelled input, can be modified to also check values
'Function type to be modified for various needs
Function InputTest() As Currency
    Dim test As String
    
    
    test = InputBox("What sales value do you want to check for?", , 0)
    If test = vbNullString Then
        MsgBox ("User canceled!")
        End
    End If
    InputTest = test

End Function
