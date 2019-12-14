Attribute VB_Name = "left_right_mid_replace"
Option Explicit

Sub stringFunctions()

    Dim myId, fName, lName, old As String
    
    myId = "Suliman Farzat is 34 years old"
    
    fName = Left(myId, 7)
    
    old = Right(myId, 12)
    
    lName = Mid(myId, 9, 6)
    
    myId = Replace(myId, "34", "20")
    
    myId = Split(myId, " ")
    
    
    
    Debug.Print vbNewLine & myId
    Debug.Print vbNewLine & fName
    Debug.Print vbNewLine & old
    Debug.Print vbNewLine & lName

End Sub
