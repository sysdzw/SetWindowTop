VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   0  'None
   Caption         =   "系统托盘管理"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   960
   End
   Begin VB.Menu mnuSys 
      Caption         =   "系统菜单"
      WindowList      =   -1  'True
      Begin VB.Menu mnuRunStartup 
         Caption         =   "设为开机自启动(&S)"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "使用帮助(&H)"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于 窗口图钉(&A)..."
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================
'名    称：窗口图钉，SetWindowTop
'描    述：窗口图钉是一款方便设置windows下窗口置顶或取消置顶的小软件。软件非常实用
'          平常都是在系统右下角托盘区，不占用任务栏。使用起来非常方便。具体使用方法
'          可右击系统托盘菜单查看帮助。
'编    程：sysdzw 原创开发，如您对本软件进行改进或拓展请发我一份
'发布日期：2020-03-17
'博    客：https://blog.csdn.net/sysdzw
'用户手册：https://www.kancloud.cn/sysdzw/clswindow/
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0 初版                                                           2020-03-17
'==============================================================================================
Option Explicit
Dim isDealing As Boolean

Private Declare Function SetForegroundWindow Lib "user32" (ByVal Hwnd As Long) As Long
 
Private Sub Form_Load()
    If App.PrevInstance Then End '防止重复运行
    
    Icon_Add Me.Hwnd, "窗口图钉", Me.Icon, 0
    mnuRunStartup.Caption = IIf(isHasSetAutoRun(), "设为手动运行(&S)", "设为开机自启动(&S)")
    
    Timer1.Enabled = True
End Sub
   
Private Sub mnuAbout_Click()
    Dim strInfo As String
    strInfo = "SetWindowTop | 窗口图钉 V" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
            "  作者:sysdzw" & vbCrLf & _
            "  主页:https://blog.csdn.net/sysdzw" & vbCrLf & _
            "  Q  Q:171977759" & vbCrLf & _
            "  邮箱:sysdzw@163.com" & vbCrLf & vbCrLf & _
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
    strHelp = "SetWindowTop | 窗口图钉 V" & App.Major & "." & App.Minor & "." & App.Revision & " 使用说明：" & vbCrLf & vbCrLf & _
        "双击启动，程序会在活动窗口的右上角显示一个控制窗口，点击下可以设置窗口置顶或者取消窗口置顶" & vbCrLf & vbCrLf & _
        "  如有问题可联系QQ171977759反馈" & vbCrLf & vbCrLf & _
        "2020-03-17"
    MsgBox strHelp, vbInformation
End Sub

Private Sub mnuRunStartup_Click()
    If mnuRunStartup.Caption = "设为手动运行(&S)" Then
        Call cancelAutoRun
        mnuRunStartup.Caption = "设为开机自启动(&S)"
    Else
        Call setAutoRun
        mnuRunStartup.Caption = "设为手动运行(&S)"
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
'创建控制窗体
Private Sub createControlWindow()
    Dim frmHdl As New frmHandle
    Load frmHdl
End Sub

Private Sub Timer1_Timer()
    Call addControlBox
End Sub
