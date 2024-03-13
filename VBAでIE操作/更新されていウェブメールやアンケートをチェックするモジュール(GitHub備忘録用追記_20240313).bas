Attribute VB_Name = "Module7"
Option Explicit

'強制的に最前面にさせる
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'元のサイズに戻す
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'public用宣言
Dim IE As Object
Dim urls As Collection

Sub run()
    'initialize
    Set IE = CreateObject("InternetExplorer.Application")
    
    '遷移urlを格納
    Set urls = New Collection
    With urls
        .Add "choose.php"
        .Add "title.php"
        .Add "open.php"
        .Add "story.php"
        .Add "agreement.php"
        .Add "enquete.php"
        .Add "column.php"
        .Add "top.do"
        .Add "finish_exec.php"
        .Add "manga.php"
    End With
    
    Call entry
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

Private Sub loading()
    'サイトの読み込み待ち
    Do While IE.Busy = True Or IE.ReadyState <> 4
        DoEvents
    Loop
    
    'JavaScript読み込み待ち
    Application.Wait Now + TimeValue("0:00:03")
End Sub

Private Sub IErefresh()
    IE.Refresh
    'JavaScript読み込み待ち
    Application.Wait Now + TimeValue("0:00:03")
End Sub

Private Sub entry()
        
    '確認ページへ
    Call IEStart("https://〇〇")
    
    'ページ更新
    Call IErefresh
    
    '既読済みでなければクリック
    Dim entryBtn As Object
    For Each entryBtn In IE.Document.getElementsByClassName("ui-btn-a")
        If "既読済み" <> entryBtn.innerText Then
            entryBtn.Click
            Call loading
            Call answer
            Exit For
        End If
    Next
    
End Sub

Private Sub answer()
    
    '呼び込み待ち ※Call loadingは待機が長すぎる
    Application.Wait Now + TimeValue("0:00:10")

    '終了ページであれば終了
    If "https://〇〇" = IE.LocationURL Then: Call toContinue
            
    'aタグ対策
    Dim linkend As Object
    If Not IsNull(IE.Document.getElementById("endlink")) Then
        IE.Document.getElementById("endlink").Click
        Call answer
    End If
    
    'form取得
    Dim form As Object, forms As Object
    Set forms = IE.Document.getElementsByTagName("form")
    
    For Each form In forms
    
        'inputタグのチェック
        Dim inputs As Object, targetinput As Object
        Set inputs = form.Document.getElementsByTagName("input")
            For Each targetinput In inputs
            
            'inputタグがラジオであればchecked
             If targetinput.Type = "radio" Then
                 targetinput.Checked = True                              'チェック
                 Exit For
             End If
        Next
        
        'チェックボックスを選択
        On Error Resume Next
        IE.Document.getElementById("que3").Checked = True
    
        'inputタグの「alt」が「進む」の場合は遷移
        If InStr(form.innerHTML, "進む") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "終了") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "コラムを読む") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "next") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "クリックしてポイントゲット") > 0 Then GoTo nextbtn
    Next

    
    For Each form In forms
        Dim i As Long
        For i = 1 To urls.count
            If InStr(form.Action, urls.Item(i)) > 0 Then
                '目的のform発見
                GoTo nextbtn
            End If
        Next i
    Next
    

    'form出ない場合、終了リンクを探す
    Dim link As Object
    For Each link In IE.Document.getElementsByTagName("a")
        If InStr(link.href, "top.do") > 0 Then
            link.Click
            Call answer
        End If
    Next

    
    Call popup
    Call answer
    Exit Sub
    
nextbtn:
    
    '送信＆再帰
    Application.Wait Now + TimeValue("0:00:03")   '数秒待つ
    
    form.submit
    Set inputs = Nothing
    Set form = Nothing
    
    Set urls = Nothing 'あとで消す
    
    Call answer
    
End Sub


Private Sub popup()
    Dim Obj As Object
    Dim re As Long
    Set Obj = CreateObject("WScript.Shell")
    re = Obj.popup("問題を起きましたデバッグします", 0, "確認", vbCritical)
    Stop
    Set Obj = Nothing
End Sub

Sub toContinue()
    Dim Obj As Object, re As Long
    Set Obj = CreateObject("WScript.Shell")
    re = Obj.popup("作業を続けますか", 0, "完了", vbYesNo + vbInformation)
    If re = vbYes Then
        IE.Quit
        Call run
    Else
        End
    End If
    Set Obj = Nothing
End Sub
