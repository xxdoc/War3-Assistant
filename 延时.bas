Attribute VB_Name = "—” ±"
Public Sub Delay(s As Long)
Dim StartTimer As Double, DelayDay As Integer
StartTimer = Timer
Do
     DoEvents
     If Timer < StartTimer Then
             DelayDay = 1
     End If
Loop Until DelayDay * 86400 + Timer - StartTimer > s / 1000
End Sub
