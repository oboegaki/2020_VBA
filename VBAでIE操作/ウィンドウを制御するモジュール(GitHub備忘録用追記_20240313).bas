Attribute VB_Name = "lib"
Option Explicit

'Dim IE As New SHDocVw.InternetExplorerMedium 'もしくは
'Dim IE As New InternetExplorerMedium 'もしくはSHDocVw.InternetExplorerMedium
'Dim IE2 As New InternetExplorerMedium

Public IE As InternetExplorerMedium
Public IE2 As InternetExplorerMedium
Public historyUrl As String


'------------------------------------------------------------------------------------------------------------
' WIN64API
'------------------------------------------------------------------------------------------------------------

'Private Declare Sub SetForegroundWindow Lib "User32" (ByVal hWnd As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'元のサイズに戻す
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'sleep
Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0                                       '' ウィンドウを非表示にし、他のウィンドウをアクティブにします。
Private Const SW_MAXIMIZE = 3                                   '' ウィンドウを最大化します。
Private Const SW_MINIMIZE = 6                                   '' ウィンドウを最小化し、Z 順位が次のトップレベルウィンドウをアクティブにします。
Private Const SW_RESTORE = 9                                    '' ウィンドウをアクティブにし、表示します。ウィンドウが最小化されていたり最大化されていたりすると、元の位置とサイズに戻ります。
Private Const SW_SHOW = 5                                       '' ウィンドウをアクティブにして、現在の位置とサイズで表示します。
Private Const SW_SHOWDEFAULT = 10                               '' アプリケーションを起動させたプログラムが CreateProcess 関数に渡すSTARTUPINFO 構造体の wShowWindow メンバで指定された SW_ フラグを基にして、表示状態を設定します。
Private Const SW_SHOWMAXIMIZED = 3                              '' ウィンドウをアクティブにして、最大化します。
Private Const SW_SHOWMINIMIZED = 2                              '' ウィンドウをアクティブにして、最小化します。
Private Const SW_SHOWMINNOACTIVE = 7                            '' ウィンドウを最小化します。アクティブなウィンドウは、アクティブな状態を維持します。非アクティブなウィンドウは、非アクティブなままです。
Private Const SW_SHOWNA = 8                                     '' ウィンドウを現在の状態で表示します。アクティブなウィンドウはアクティブな状態を維持します。
Private Const SW_SHOWNOACTIVATE = 4                             '' ウィンドウを直前の位置とサイズで表示します。アクティブなウィンドウはアクティブな状態を維持します。
Private Const SW_SHOWNORMAL = 1                                 '' ウィンドウをアクティブにして、表示します。ウィンドウが最小化または最大化されているときは、位置とサイズを元に戻します。


'------------------------------------------------------------------------------------------------------------
' メッセージボックス
'------------------------------------------------------------------------------------------------------------
'最前面に表示(秒数指定可能)
'WSH.Popup(strText,[nSecondsToWait],[strTitle],[nType]) 1.メッセージ 2.秒数 3.タイトル 4.アイコンやボタンの種類
Sub msgpop(msg As String, Optional time As Long = 0, Optional title As String = "お知らせ", Optional context As Variant = vbInformation)
    Dim re As Long
    re = CreateObject("WScript.Shell").popup(msg, 0, title, context)
    
    '最小化されている場合は元に戻す(9=RESTORE:最小化前の状態)
    If IsIconic(re) Then
        ShowWindowAsync re, &H9
    End If
    SetForegroundWindow (re)     '最前面に表示
    'With CreateObject("WScript.Shell").popup(msg, 0, title, context): End With
End Sub

'------------------------------------------------------------------------------------------------------------
' IEライブラリ
'------------------------------------------------------------------------------------------------------------

Sub IE_OPEN(targetURL As String, IE As InternetExplorerMedium)
    'オートメーションエラー対策
    '参考①: https://qiita.com/3mc/items/da045e86d25ef697ec43
    '参考②: https://stackoverflow.com/questions/12965032/excel-vba-controlling-ie-local-intranet
    
    'エラー対策
    historyUrl = targetURL
    
    Dim objIE As InternetExplorerMedium
    Set objIE = New InternetExplorerMedium ' This object (the "medium" variety as opposed to "InternetExplorer") is necessary in our security climate
    'Dim targetURL As String
    Dim webContent As String
    Dim sh
    Dim eachIE
    
    '   最小化されている場合は元に戻す(9=RESTORE:最小化前の状態)
    '    If IsIconic(IE.hWnd) Then
    '        ShowWindowAsync IE.hWnd, &H9
    '    End If
    '    SetForegroundWindow (IE.hWnd) '最前面に表示
    
    objIE.Visible = False ' Set to true to watch what's happening
    objIE.Navigate targetURL
    
    
    '最大化最小化アクティなど
    '参考; http://www.asahi-net.or.jp/~fq7y-krsk/vba_v1.2/Win32API_VBA/xShowWindow.html
    'ShowWindowAsync objIE.hwnd, 6 '最小化してトップレベルのウィンドウをアクティブにする
    
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


'読み込み待ち
Sub IE_Wait(IE As InternetExplorerMedium)

    'IEがない場合は再読み込み
    If IE Is Nothing Then
        IE_OPEN historyUrl, IE
    End If

    While IE.Busy  ' The new process may still be busy even after you find it
      DoEvents: Sleep 1
    Wend
    
    '最大化最小化アクティなど → http://www.asahi-net.or.jp/~fq7y-krsk/vba_v1.2/Win32API_VBA/xShowWindow.html
    'ShowWindowAsync IE.hwnd, 2 'アクティブにして最小化
    

    'IE.Visible = True
    Dim waitMax As Long: waitMax = 0
    Do While IE.Busy = True Or IE.ReadyState <> 4
        DoEvents: Sleep 1
        Application.Wait Now + TimeValue("0:00:01")
        waitMax = waitMax + 1
        If waitMax > 30 Then Exit Do
    Loop
    

'    '最小化されている場合は元に戻す(9=RESTORE:最小化前の状態)
'    If IsIconic(IE.hWnd) Then
'        ShowWindowAsync IE.hWnd, &H9
'    End If
'    SetForegroundWindow (IE.hWnd) '最前面に表示

'    最大化最小化アクティなど → http://www.asahi-net.or.jp/~fq7y-krsk/vba_v1.2/Win32API_VBA/xShowWindow.html
'    ShowWindowAsync IE.hwnd, 2 'アクティブにして最小化
    
    '最大化
'    Dim ret As Long
'    ret = ShowWindowAsync(IE.hWnd, 3)
    
    'JavaScript読み込み待ち
    Sleep 3000 'Application.Wait Now + TimeValue("0:00:03")
    
    'ページエラー対策
    If InStr(IE.Document.title, "このページを表示できません") > 0 Or InStr(IE.Document.title, "安全では") > 0 Then
        'MsgBox "エラー発生"
        'IE.Document.getElementById("task1-3").Click
        
        IE.Refresh
        Sleep (15000)
                
        If InStr(IE.Document.title, "このページを表示できません") > 0 Or InStr(IE.Document.title, "安全では") > 0 Then
            SendKeys "{F5}"
        End If
        
    End If
    
End Sub

' IEの画像表示・非表示切り替え（レジストリ）
Sub IE_ImageOn()
    Dim ret As Variant: ret = IsIE_Image(True)  '表示
End Sub
Sub IE_ImageOff()
    Dim ret As Variant: ret = IsIE_Image(False)  '非表示
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
    Application.Cursor = xlDefault 'カーソルを元に戻す
End Sub

'タスク上のすべてのIE窓を閉じる
Sub AllcloseIEWindow()
    Application.Wait Now + TimeValue("0:00:05")
    Dim win As Object
    For Each win In CreateObject("Shell.Application").Windows
        If win.Name = "Internet Explorer" And win.Visible = True Then
             win.Quit
        End If
    Next
End Sub

'指定のIE以外のIEWindowsを閉じる
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

'別窓のIEをセット
'参考
'  ①https://vba-code.net/ie/open-link-in-new-tab
'  ②http://kouten0430.hatenablog.com/entry/2018/08/11/150052
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
    Set WD = CreateObject("Word.Application")    ''Wordを起動します
    For Each task In WD.Tasks                    ''Word VBAのTasksコレクションを調べます
        If task.Visible = True Then              ''タスク(プロセス)が実行中だったら
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

'※途中からやり直すためのプロシージャ
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



