Attribute VB_Name = "modPub"
Option Explicit
Public lngHandleHwnd As Long
Sub Main()
    If App.PrevInstance Then End '��ֹ�ظ�����

    Load frmStartup
    
    Dim w As New clsWindow, s$, s2$, v, i%
    w.GetWindowByClassNameEx "", 0, s, True, , DisplayedWindow
    v = Split(s, " ")
    For i = 0 To UBound(v)
        If v(i) <> Me.hwnd Then
            w.hwnd = v(i)
            lngHandleHwnd = w.hwnd
            Call createControlWindow
            s2 = s2 & i & " " & v(i) & "(" & w.Width & "," & w.Height & ")" & w.Top & " " & w.Caption & vbCrLf
        End If
    Next
End Sub
'�������ƴ���
Private Sub createControlWindow()
    Dim frmHdl As New frmHandle
    Load frmHdl
End Sub
