VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   0  'None
   Caption         =   "ϵͳ���̹���"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   960
   End
   Begin VB.Menu mnuSys 
      Caption         =   "ϵͳ�˵�"
      WindowList      =   -1  'True
      Begin VB.Menu mnuRunStartup 
         Caption         =   "��Ϊ����������(&S)"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "ʹ�ð���(&H)"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "���� ����ͼ��(&A)..."
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================
'��    �ƣ�����ͼ����SetWindowTop
'��    ��������ͼ����һ�������windows�´����ö���ȡ���ö���С���������ǳ�ʵ��
'          ƽ��������ϵͳ���½�����������ռ����������ʹ�������ǳ����㡣����ʹ�÷���
'          ���һ�ϵͳ���̲˵��鿴������
'��    �̣�sysdzw ԭ�������������Ա�������иĽ�����չ�뷢��һ��
'�������ڣ�2020-03-17
'��    �ͣ�https://blog.csdn.net/sysdzw
'�û��ֲ᣺https://www.kancloud.cn/sysdzw/clswindow/
'Email   ��sysdzw@163.com
'QQ      ��171977759
'��    ����V1.0 ����                                                           2020-03-17
'==============================================================================================
Option Explicit
Dim isDealing As Boolean

Private Declare Function SetForegroundWindow Lib "user32" (ByVal Hwnd As Long) As Long
 
Private Sub Form_Load()
    If App.PrevInstance Then End '��ֹ�ظ�����
    
    Icon_Add Me.Hwnd, "����ͼ��", Me.Icon, 0
    mnuRunStartup.Caption = IIf(isHasSetAutoRun(), "��Ϊ�ֶ�����(&S)", "��Ϊ����������(&S)")
    
    Timer1.Enabled = True
End Sub
   
Private Sub mnuAbout_Click()
    Dim strInfo As String
    strInfo = "SetWindowTop | ����ͼ�� V" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
            "  ����:sysdzw" & vbCrLf & _
            "  ��ҳ:https://blog.csdn.net/sysdzw" & vbCrLf & _
            "  Q  Q:171977759" & vbCrLf & _
            "  ����:sysdzw@163.com" & vbCrLf & vbCrLf & _
            "2020-03-17"
    MsgBox strInfo, vbInformation
End Sub

Public Sub mnuExit_Click()
    Call Icon_Del(Me.Hwnd, 0)
    
    Dim frm As Form
    Dim w As New clsWindow
    For Each frm In Forms
        If frm.Caption = "SetWindowTop" Then
            Unload frm
        End If
    Next
    
    Unload Me
End Sub

Private Sub mnuHelp_Click()
    Dim strHelp$
    strHelp = "SetWindowTop | ����ͼ�� V" & App.Major & "." & App.Minor & "." & App.Revision & " ʹ��˵����" & vbCrLf & vbCrLf & _
        "˫��������������ڻ���ڵ����Ͻ���ʾһ�����ƴ��ڣ�����¿������ô����ö�����ȡ�������ö�" & vbCrLf & vbCrLf & _
        "  �����������ϵQQ171977759����" & vbCrLf & vbCrLf & _
        "2020-03-17"
    MsgBox strHelp, vbInformation
End Sub

Private Sub mnuRunStartup_Click()
    If mnuRunStartup.Caption = "��Ϊ�ֶ�����(&S)" Then
        Call cancelAutoRun
        mnuRunStartup.Caption = "��Ϊ����������(&S)"
    Else
        Call setAutoRun
        mnuRunStartup.Caption = "��Ϊ�ֶ�����(&S)"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX

    lMsg = X / Screen.TwipsPerPixelX
    Select Case lMsg
    Case WM_RBUTTONUP
        SetForegroundWindow Me.Hwnd
        PopupMenu mnuSys
    End Select
End Sub
Private Sub addControlBox()
    Dim w As New clsWindow, s$, v, i%
    w.GetWindowByTitleEx ".+?", 0, s, True, , DisplayedWindow
    v = Split(s, " ")
    For i = 0 To UBound(v)
        If v(i) <> Me.Hwnd Then
            If Not isHasAddControlBox(v(i)) Then
                w.Hwnd = v(i)
                If InStr("|SetWindowTop|Program Manager|", "|" & w.Caption & "|") = 0 Then
                    lngHandleHwnd = w.Hwnd
                    Call createControlWindow
                End If
            End If
        End If
    Next
End Sub

Private Function isHasAddControlBox(ByVal lngHwnd As Long) As Boolean
    Dim frm As Form
    Dim w As New clsWindow
    For Each frm In Forms
        If frm.Caption = "SetWindowTop" And frm.Tag = CStr(lngHwnd) Then
            isHasAddControlBox = True
            Exit Function
        End If
    Next
End Function
'�������ƴ���
Private Sub createControlWindow()
    Dim frmHdl As New frmHandle
    Load frmHdl
End Sub

Private Sub Timer1_Timer()
    Call addControlBox
End Sub
