Attribute VB_Name = "��ʼ��"
Sub Main()
Dim RichOCX() As Byte, Lon As Long
If Dir("C:\WINDOWS\system32\RICHTX32.OCX") = "" Then    '���RichText�ؼ�������
        RichOCX = LoadResData(106, "CUSTOM")
        Open "C:\WINDOWS\system32\RICHTX32.OCX" For Binary As #1
        For Lon = 0 To UBound(RichOCX)  '�����ɵ��ļ���С
        Put #1, , RichOCX(Lon) '�ͷ��ļ�
        Next
        Close #1
        Shell "regsvr32 /s C:\WINDOWS\system32\RICHTX32.OCX", 0
End If

If Dir("C:\WINDOWS\system32\comdlg32.ocx") = "" Then    '���comdlg32k�ؼ�������
    RichOCX = LoadResData(108, "CUSTOM")
    Open "C:\WINDOWS\system32\comdlg32.ocx" For Binary As #1
    For Lon = 0 To UBound(RichOCX)  '�����ɵ��ļ���С
        Put #1, , RichOCX(Lon) '�ͷ��ļ�
    Next
        Close #1
        Shell "regsvr32 /s C:\WINDOWS\system32\comdlg32.ocx", 0
End If
Form1.Show
End Sub
