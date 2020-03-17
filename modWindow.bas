Attribute VB_Name = "modWindow"
'=====================================================================================
'��    ������clsWindow.cls�������ģ�飬һЩ�޷��ŵ���ģ���еĴ���������� (modWindow)
'��    �̣�sysdzw ԭ���������������Ҫ��ģ����и����뷢��һ�ݣ���ͬά��
'�������ڣ�2013/05/28
'��    �ͣ�http://blog.csdn.net/sysdzw
'�û��ֲ᣺https://www.kancloud.cn/sysdzw/clswindow/
'Email   ��sysdzw@163.com
'QQ      ��171977759
'��    ����V1.0 ����                                        2012/12/3
'          V1.1 �����е�api�����Լ����ֱ���Ų����ģ��         2013/05/28
'          V1.2 ��EnumChildProc�л�ȡ�ؼ����ֺ����޸���      2013/06/13
'          V1.3 ����ģ�������Ƶ���ģ���еĶ��ƹ�ȥ��          2020/01/19
'               ��GetText�����ŵ��������                   2020/03/12
'=====================================================================================
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const lMaxLength& = 500
Public strControlInfo$ '�������������пؼ�����Ϣ
Public lngHandleHwnd As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ���api����EnumChildWindows���ʹ�õõ�һ�����������ڵ�����child�ؼ�
'��������EnumChildProc
'��ڲ�����hWnd   long��  ���������һ��ָ������
'����ֵ��long   ����ֱ�ӷ��ص�true�������true���������
'��ע��sysdzw �� 2010-11-13 �ṩ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim strClassName As String * 256
    Dim strText As String
    Dim lngCtlId As Long
    Dim strHwnd$, strCtlId$, strClass$, lRet&

    EnumChildProc = True

    lngCtlId = GetWindowLong(hwnd, (-12))
    lRet = GetClassName(hwnd, strClassName, 255)

    strText = GetText(hwnd)
    strText = Replace(strText, vbCrLf, " ") 'ǿ�ƽ��ı��������ݻس��滻�ɿո��Է�ֹӰ�������ȡ

    strHwnd$ = CStr(hwnd) & vbTab
    strCtlId$ = CStr(lngCtlId) & vbTab
    strClass$ = Left$(strClassName, lRet) & vbTab
    strControlInfo = strControlInfo & strHwnd$ & _
                    strCtlId$ & _
                    strClass$ & _
                    strText & vbCrLf
End Function
'���ݾ����ô�������
Public Function GetText(ByVal hwnd As Long) As String
    '����1 ����һ��
'    Dim Txt2() As Byte, i&
'    i = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
'    If i = 0 Then Exit Function 'û������
'    ReDim Txt2(i)
'    SendMessage hWnd, WM_GETTEXT, i + 1, Txt2(0)
'    ReDim Preserve Txt2(i - 1)
'    GetTextByHwnd = StrConv(Txt2, vbUnicode)

    '����2 ��Ϸ�������������api���ã�������������С���ṩ��
    Dim Txt2() As Byte, i&
    ReDim Txt2(lMaxLength&) '���ʵ�����ݶ���һ���ֽ���װ������0
    SendMessage hwnd, &HD, lMaxLength&, Txt2(0)
    If Txt2(0) = 0 Then Exit Function  'û������
    For i = 1 To lMaxLength&
        If Txt2(i) = 0 Then Exit For '����
    Next
    If i >= lMaxLength - 2& Then '����ӽ�����Ϊȡ���ݲ�������ֱ����api���㳤��ȡ
        i = SendMessage(hwnd, &HE, 0&, 0&)
        If i = 0 Then Exit Function 'û������
        ReDim Txt2(i) '���ʵ�����ݶ���һ���ֽ���װ������0
        SendMessage hwnd, &HD, i + 1, Txt2(0)
    End If
    ReDim Preserve Txt2(i - 1) 'ȥ������ֽ�
    GetText = StrConv(Txt2, vbUnicode) 'תASI�ִ�Ϊ���ִ�
End Function
