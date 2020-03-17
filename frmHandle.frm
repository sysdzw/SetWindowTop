VERSION 5.00
Begin VB.Form frmHandle 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "SetWindowTop"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2370
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   840
   End
   Begin VB.Image imgBefore 
      Height          =   255
      Left            =   1080
      Picture         =   "frmHandle.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgAfter 
      Height          =   255
      Left            =   1440
      Picture         =   "frmHandle.frx":015F
      Stretch         =   -1  'True
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmHandle.frx":02AF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wTag As New clsWindow
Dim wMe As New clsWindow

Dim lngLeft As Long
Dim lngTop As Long
Dim isClickHandle As Boolean
Dim isTop As Boolean

Private Sub Form_Load()
    Me.Tag = lngHandleHwnd
    wTag.Hwnd = lngHandleHwnd
    Timer1.Enabled = True

    isTop = wTag.IsTopmost
    Set Image1.Picture = IIf(isTop, imgAfter.Picture, imgBefore.Picture)

    Me.Width = 255
    Me.Height = 255
    
    wMe.Hwnd = Me.Hwnd
    wMe.Transparent 50
    wMe.SetTop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    setTagPos
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    setTagPos
End Sub

Private Sub setTagPos()
    isClickHandle = True
    wTag.SetTop Not wTag.IsTopmost
    
    Set Image1.Picture = IIf(wTag.IsTopmost, imgAfter.Picture, imgBefore.Picture)
    If wTag.IsTopmost Then
        wTag.SetTop
        wTag.Focus
        wMe.SetTop
    End If
    wTag.Focus
End Sub

Private Sub Timer1_Timer()
    If Not wTag.CheckWindow Then '如果窗口不存在就关闭
        Unload Me
    End If
    
    If wTag.IsForegroundWindow Then '如果当前是活动窗口，那么需要显示并移动控制窗口
        lngLeft = (wTag.Left + wTag.Width) * 15 - Me.Width - 60 * 15
        lngTop = wTag.Top * 15 + 60
        If Me.Left <> lngLeft Or Me.Top <> lngTop Then '位置需要更新时再移动
            Me.Move lngLeft, lngTop
        End If
        Me.Visible = True
        wMe.SetTop
    ElseIf Not wMe.IsForegroundWindow Then
        Me.Visible = False
    End If
    
    If wTag.IsTopmost <> isTop Then
        isTop = Not isTop
        Set Image1.Picture = IIf(isTop, imgAfter.Picture, imgBefore.Picture)
    End If
End Sub
