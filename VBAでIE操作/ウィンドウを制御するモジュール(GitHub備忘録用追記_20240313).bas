Attribute VB_Name = "lib"
Option Explicit

'Dim IE As New SHDocVw.InternetExplorerMedium '��������
'Dim IE As New InternetExplorerMedium '��������SHDocVw.InternetExplorerMedium
'Dim IE2 As New InternetExplorerMedium

Public IE As InternetExplorerMedium
Public IE2 As InternetExplorerMedium
Public historyUrl As String


'------------------------------------------------------------------------------------------------------------
' WIN64API
'------------------------------------------------------------------------------------------------------------

'Private Declare Sub SetForegroundWindow Lib "User32" (ByVal hWnd As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'���̃T�C�Y�ɖ߂�
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'sleep
Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0                                       '' �E�B���h�E���\���ɂ��A���̃E�B���h�E���A�N�e�B�u�ɂ��܂��B
Private Const SW_MAXIMIZE = 3                                   '' �E�B���h�E���ő剻���܂��B
Private Const SW_MINIMIZE = 6                                   '' �E�B���h�E���ŏ������AZ ���ʂ����̃g�b�v���x���E�B���h�E���A�N�e�B�u�ɂ��܂��B
Private Const SW_RESTORE = 9                                    '' �E�B���h�E���A�N�e�B�u�ɂ��A�\�����܂��B�E�B���h�E���ŏ�������Ă�����ő剻����Ă����肷��ƁA���̈ʒu�ƃT�C�Y�ɖ߂�܂��B
Private Const SW_SHOW = 5                                       '' �E�B���h�E���A�N�e�B�u�ɂ��āA���݂̈ʒu�ƃT�C�Y�ŕ\�����܂��B
Private Const SW_SHOWDEFAULT = 10                               '' �A�v���P�[�V�������N���������v���O������ CreateProcess �֐��ɓn��STARTUPINFO �\���̂� wShowWindow �����o�Ŏw�肳�ꂽ SW_ �t���O����ɂ��āA�\����Ԃ�ݒ肵�܂��B
Private Const SW_SHOWMAXIMIZED = 3                              '' �E�B���h�E���A�N�e�B�u�ɂ��āA�ő剻���܂��B
Private Const SW_SHOWMINIMIZED = 2                              '' �E�B���h�E���A�N�e�B�u�ɂ��āA�ŏ������܂��B
Private Const SW_SHOWMINNOACTIVE = 7                            '' �E�B���h�E���ŏ������܂��B�A�N�e�B�u�ȃE�B���h�E�́A�A�N�e�B�u�ȏ�Ԃ��ێ����܂��B��A�N�e�B�u�ȃE�B���h�E�́A��A�N�e�B�u�Ȃ܂܂ł��B
Private Const SW_SHOWNA = 8                                     '' �E�B���h�E�����݂̏�Ԃŕ\�����܂��B�A�N�e�B�u�ȃE�B���h�E�̓A�N�e�B�u�ȏ�Ԃ��ێ����܂��B
Private Const SW_SHOWNOACTIVATE = 4                             '' �E�B���h�E�𒼑O�̈ʒu�ƃT�C�Y�ŕ\�����܂��B�A�N�e�B�u�ȃE�B���h�E�̓A�N�e�B�u�ȏ�Ԃ��ێ����܂��B
Private Const SW_SHOWNORMAL = 1                                 '' �E�B���h�E���A�N�e�B�u�ɂ��āA�\�����܂��B�E�B���h�E���ŏ����܂��͍ő剻����Ă���Ƃ��́A�ʒu�ƃT�C�Y�����ɖ߂��܂��B


'------------------------------------------------------------------------------------------------------------
' ���b�Z�[�W�{�b�N�X
'------------------------------------------------------------------------------------------------------------
'�őO�ʂɕ\��(�b���w��\)
'WSH.Popup(strText,[nSecondsToWait],[strTitle],[nType]) 1.���b�Z�[�W 2.�b�� 3.�^�C�g�� 4.�A�C�R����{�^���̎��
Sub msgpop(msg As String, Optional time As Long = 0, Optional title As String = "���m�点", Optional context As Variant = vbInformation)
    Dim re As Long
    re = CreateObject("WScript.Shell").popup(msg, 0, title, context)
    
    '�ŏ�������Ă���ꍇ�͌��ɖ߂�(9=RESTORE:�ŏ����O�̏��)
    If IsIconic(re) Then
        ShowWindowAsync re, &H9
    End If
    SetForegroundWindow (re)     '�őO�ʂɕ\��
    'With CreateObject("WScript.Shell").popup(msg, 0, title, context): End With
End Sub

'------------------------------------------------------------------------------------------------------------
' IE���C�u����
'------------------------------------------------------------------------------------------------------------

Sub IE_OPEN(targetURL As String, IE As InternetExplorerMedium)
    '�I�[�g���[�V�����G���[�΍�
    '�Q�l�@: https://qiita.com/3mc/items/da045e86d25ef697ec43
    '�Q�l�A: https://stackoverflow.com/questions/12965032/excel-vba-controlling-ie-local-intranet
    
    '�G���[�΍�
    historyUrl = targetURL
    
    Dim objIE As InternetExplorerMedium
    Set objIE = New InternetExplorerMedium ' This object (the "medium" variety as opposed to "InternetExplorer") is necessary in our security climate
    'Dim targetURL As String
    Dim webContent As String
    Dim sh
    Dim eachIE
    
    '   �ŏ�������Ă���ꍇ�͌��ɖ߂�(9=RESTORE:�ŏ����O�̏��)
    '    If IsIconic(IE.hWnd) Then
    '        ShowWindowAsync IE.hWnd, &H9
    '    End If
    '    SetForegroundWindow (IE.hWnd) '�őO�ʂɕ\��
    
    objIE.Visible = False ' Set to true to watch what's happening
    objIE.Navigate targetURL
    
    
    '�ő剻�ŏ����A�N�e�B�Ȃ�
    '�Q�l; http://www.asahi-net.or.jp/~fq7y-krsk/vba_v1.2/Win32API_VBA/xShowWindow.html
    'ShowWindowAsync objIE.hwnd, 6 '�ŏ������ăg�b�v���x���̃E�B���h�E���A�N�e�B�u�ɂ���
    
    Sleep (6000)
    'While objIE.Busy: DoEvents: Wend
      
    Do
      'Set sh = New Shell32.Shell
      'For Each eachIE In sh.Windows
      For Each eachIE In CreateObject("Shell.Application").Windows
          DoEvents: Sleep 1
          If InStr(1, eachIE.LocationURL, targetURL) Then
            Set IE = eachIE
          'IE.Visible = False  'This is here because in some environments, the new process defaults to Visible.
          Exit Do
          End If
        Next eachIE
      Loop
    Set eachIE = Nothing
    Set sh = Nothing
    Set objIE = Nothing
    
    While IE.Busy
      DoEvents: Sleep 1
      Wend
      
    IE_Wait IE

End Sub


'�ǂݍ��ݑ҂�
Sub IE_Wait(IE As InternetExplorerMedium)

    'IE���Ȃ��ꍇ�͍ēǂݍ���
    If IE Is Nothing Then
        IE_OPEN historyUrl, IE
    End If

    While IE.Busy  ' The new process may still be busy even after you find it
      DoEvents: Sleep 1
    Wend
    
    '�ő剻�ŏ����A�N�e�B�Ȃ� �� http://www.asahi-net.or.jp/~fq7y-krsk/vba_v1.2/Win32API_VBA/xShowWindow.html
    'ShowWindowAsync IE.hwnd, 2 '�A�N�e�B�u�ɂ��čŏ���
    

    'IE.Visible = True
    Dim waitMax As Long: waitMax = 0
    Do While IE.Busy = True Or IE.ReadyState <> 4
        DoEvents: Sleep 1
        Application.Wait Now + TimeValue("0:00:01")
        waitMax = waitMax + 1
        If waitMax > 30 Then Exit Do
    Loop
    

'    '�ŏ�������Ă���ꍇ�͌��ɖ߂�(9=RESTORE:�ŏ����O�̏��)
'    If IsIconic(IE.hWnd) Then
'        ShowWindowAsync IE.hWnd, &H9
'    End If
'    SetForegroundWindow (IE.hWnd) '�őO�ʂɕ\��

'    �ő剻�ŏ����A�N�e�B�Ȃ� �� http://www.asahi-net.or.jp/~fq7y-krsk/vba_v1.2/Win32API_VBA/xShowWindow.html
'    ShowWindowAsync IE.hwnd, 2 '�A�N�e�B�u�ɂ��čŏ���
    
    '�ő剻
'    Dim ret As Long
'    ret = ShowWindowAsync(IE.hWnd, 3)
    
    'JavaScript�ǂݍ��ݑ҂�
    Sleep 3000 'Application.Wait Now + TimeValue("0:00:03")
    
    '�y�[�W�G���[�΍�
    If InStr(IE.Document.title, "���̃y�[�W��\���ł��܂���") > 0 Or InStr(IE.Document.title, "���S�ł�") > 0 Then
        'MsgBox "�G���[����"
        'IE.Document.getElementById("task1-3").Click
        
        IE.Refresh
        Sleep (15000)
                
        If InStr(IE.Document.title, "���̃y�[�W��\���ł��܂���") > 0 Or InStr(IE.Document.title, "���S�ł�") > 0 Then
            SendKeys "{F5}"
        End If
        
    End If
    
End Sub

' IE�̉摜�\���E��\���؂�ւ��i���W�X�g���j
Sub IE_ImageOn()
    Dim ret As Variant: ret = IsIE_Image(True)  '�\��
End Sub
Sub IE_ImageOff()
    Dim ret As Variant: ret = IsIE_Image(False)  '��\��
End Sub
Private Function IsIE_Image(Optional flg As Boolean)
    Const IE_MAGE As String = "HKCU\Software\Microsoft\Internet Explorer\Main\"
    Const DISPLAY_INL As String = "Display Inline Images"
    Dim wsh As Object
    Dim RegValue
    Set wsh = CreateObject("WScript.Shell")
    RegValue = wsh.RegRead(IE_MAGE & DISPLAY_INL)
    
    If flg Then
        wsh.RegWrite IE_MAGE & DISPLAY_INL, "yes", "REG_SZ"
    Else
        wsh.RegWrite IE_MAGE & DISPLAY_INL, "no", "REG_SZ"
    End If
    
    RegValue = ""
    IsIE_Image = wsh.RegRead(IE_MAGE & DISPLAY_INL)
End Function

Sub CursorDefault()
    Application.Cursor = xlDefault '�J�[�\�������ɖ߂�
End Sub

'�^�X�N��̂��ׂĂ�IE�������
Sub AllcloseIEWindow()
    Application.Wait Now + TimeValue("0:00:05")
    Dim win As Object
    For Each win In CreateObject("Shell.Application").Windows
        If win.Name = "Internet Explorer" And win.Visible = True Then
             win.Quit
        End If
    Next
End Sub

'�w���IE�ȊO��IEWindows�����
Sub IE_closeAnotherWins(IE As InternetExplorerMedium)
    Dim win As Object
    For Each win In CreateObject("Shell.Application").Windows
        If win.Name = "Internet Explorer" Then
            If IE.hWnd <> win.hWnd And win.Visible = True Then
                win.Quit
            End If
        End If
    Next
End Sub

'�ʑ���IE���Z�b�g
'�Q�l
'  �@https://vba-code.net/ie/open-link-in-new-tab
'  �Ahttp://kouten0430.hatenablog.com/entry/2018/08/11/150052
Function IE_OptChangeWin(IE As IWebBrowser2) As IWebBrowser2
    IE.Visible = False
    Sleep (3000)
    Dim win As Object, IE2 As IWebBrowser2
    For Each win In CreateObject("Shell.Application").Windows
        If win.Name = "Internet Explorer" Then
            If win.Visible = True Then
                Set IE2 = win
                Exit For
            End If
        End If
    Next
    IE.Visible = True
End Function


Sub Sample1()
    Dim WD, task, n As Long
    Set WD = CreateObject("Word.Application")    ''Word���N�����܂�
    For Each task In WD.Tasks                    ''Word VBA��Tasks�R���N�V�����𒲂ׂ܂�
        If task.Visible = True Then              ''�^�X�N(�v���Z�X)�����s����������
            Debug.Print (task.Name)
        End If
    Next
    WD.Quit
    Set WD = Nothing
End Sub


Sub SetOnlyWindow()
    Dim win As Object
    For Each win In CreateObject("Shell.Application").Windows
        If win.Name = "Internet Explorer" Then
            If win.Visible = True Then
                Set IE = win
                Exit For
            End If
        End If
    Next
End Sub

'���r�������蒼�����߂̃v���V�[�W��
Function getIE() As IWebBrowser2
    Dim win As Object
    For Each win In CreateObject("Shell.Application").Windows
        If win.Name = "Internet Explorer" Then
            If win.Visible = True Then
                Set getIE = win
                Exit For
            End If
        End If
    Next
End Function



