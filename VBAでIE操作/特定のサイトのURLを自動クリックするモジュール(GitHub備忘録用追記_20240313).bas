Attribute VB_Name = "Enavi"
Option Explicit


'強制的に最前面にさせる
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'元のサイズに戻す
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'public用宣言
Dim IE As Object

Sub run()
    'initialize
    Set IE = CreateObject("InternetExplorer.Application")
    
    'ログインページ遷移
    Call IEStart("https://〇〇")
    Call login
    
End Sub

'----------------------------------------------------------
'STEP1 : IE起動
'----------------------------------------------------------
Private Sub IEStart(ByVal url As String)
    IE.Visible = True
    IE.Navigate url

    '最小化されている場合は元に戻す(9=RESTORE:最小化前の状態)
    If IsIconic(IE.hWnd) Then
        ShowWindowAsync IE.hWnd, &H9
    End If
    SetForegroundWindow (IE.hWnd)     '最前面に表示
    
    Call loading
End Sub


'----------------------------------------------------------
'STEP2 : ログイン
'----------------------------------------------------------
Private Sub login()

    'ID・PASS入力
    IE.Document.getElementById("u").Value = "ユーザ名"
    IE.Document.getElementById("p").Value = "パスワード"

    '4秒待機
    'WScript.Sleep 5000
    Application.Wait Now + TimeValue("0:00:04")
    
    'ログイン
    IE.Document.getElementById("loginButton").Click
    Call loading
    
    'Enaviトップへ
    Call linksClick

End Sub


'----------------------------------------------------------
'STEP3 : ログイン後にリンクをクリックしたいページに遷移
'----------------------------------------------------------
Private Sub linksClick()
    
    Call IEStart("https://〇〇")
    
    'IEオブジェクト内のリンクをチェックし、クリックしたいリンクを探す
     Dim i As Long
     For i = 0 To IE.Document.links.Length - 1
       '取得したリンク先アドレスを文字列単位で比較
       If IE.Document.links(i).href = "javascript:void(0);" Then
          'クリック対象だったらクリックしてループを抜ける
           IE.Document.links(i).Click
           Exit For
       End If
    Next i
    
    Call loading
        
    Application.Wait Now + TimeValue("0:00:10")
    
    IE.Quit
    Set IE = Nothing
End Sub


Private Sub loading()
    'サイトの読み込み待ち
    Do While IE.Busy = True Or IE.ReadyState <> 4
        DoEvents
    Loop
    'JavaScript読み込み待ち
    Application.Wait Now + TimeValue("0:00:03")
End Sub


