Attribute VB_Name = "dataType"
Option Explicit

Sub dataTyp()

    Dim bo As Boolean
    Dim bt As Byte
    Dim cr As Currency
    Dim dt As Date
    Dim db As Double
    Dim ig As Integer
    Dim lg As Long
    Dim st As String
    
    bo = True
    bt = 200
    cr = -100
    dt = "02/11/2019"
    db = 1000000.12345
    ig = 30000
    lg = -2147483647
    st = "String"" (fixed-length)"
    
    Range("A1") = bo
    Range("A2") = bt
    Range("A3") = cr
    Range("A4") = dt
    Range("A5") = db
    Range("A6") = ig
    Range("A7") = lg
    Range("A8") = st
    
End Sub

Sub whatData()

Range("B1") = TypeName(Range("A1").Value)
Range("B2") = TypeName(Range("A2").Value)
Range("B3") = TypeName(Range("A3").Value)
Range("B4") = TypeName(Range("A4").Value)
Range("B5") = TypeName(Range("A5").Value)
Range("B6") = TypeName(Range("A6").Value)
Range("B7") = TypeName(Range("A7").Value)
Range("B8") = TypeName(Range("A8").Value)

End Sub
