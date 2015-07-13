Attribute VB_Name = "初始化"
Sub Main()
Dim RichOCX() As Byte, Lon As Long
If Dir("C:\WINDOWS\system32\RICHTX32.OCX") = "" Then    '如果RichText控件不存在
        RichOCX = LoadResData(106, "CUSTOM")
        Open "C:\WINDOWS\system32\RICHTX32.OCX" For Binary As #1
        For Lon = 0 To UBound(RichOCX)  '欲生成的文件大小
        Put #1, , RichOCX(Lon) '释放文件
        Next
        Close #1
        Shell "regsvr32 /s C:\WINDOWS\system32\RICHTX32.OCX", 0
End If

If Dir("C:\WINDOWS\system32\comdlg32.ocx") = "" Then    '如果comdlg32k控件不存在
    RichOCX = LoadResData(108, "CUSTOM")
    Open "C:\WINDOWS\system32\comdlg32.ocx" For Binary As #1
    For Lon = 0 To UBound(RichOCX)  '欲生成的文件大小
        Put #1, , RichOCX(Lon) '释放文件
    Next
        Close #1
        Shell "regsvr32 /s C:\WINDOWS\system32\comdlg32.ocx", 0
End If
Form1.Show
End Sub
