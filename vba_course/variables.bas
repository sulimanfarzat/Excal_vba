Attribute VB_Name = "variables"
Option Explicit


Sub variables()

    Dim str As String: str = "how are you"
    Dim num1, num2 As Double
    Dim result As Double
    Dim today As Date
    Const year As Integer = 365
    Dim isOk As Boolean
    
    
    str = "Hello world"
    num1 = 5
    num2 = 0 + 0.00001
    result = num1 / num2
    today = #3/10/2019#
    isOk = True
    
    
    Debug.Print str
    Debug.Print "#######################"
    Debug.Print result
    Debug.Print "#######################"
    Debug.Print today
    Debug.Print "#######################"
    Debug.Print year
    Debug.Print "#######################"
    Debug.Print isOk

End Sub













Sub variables_2()

    Debug.Print "sub_var_2 " & str

End Sub
