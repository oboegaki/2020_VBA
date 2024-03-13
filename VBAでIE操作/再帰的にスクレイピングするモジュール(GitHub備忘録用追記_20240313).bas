Attribute VB_Name = "S002_mail_de_point"
Option Explicit


Sub run()
    Application.StatusBar = "起動中..."
    lib.IE_ImageOff

    Call Top
End Sub

Private Sub Top()
        
    'visibleなIEを全て閉じる
    Call lib.AllcloseIEWindow
    lib.IE_OPEN "https://〇〇", IE
    
    'javascriptが出るまで長めに待機
    Application.Wait Now + TimeValue("0:00:15")
    
    'チェック
    'Dim targetlist: Set targetlist = IE.Document.getElementsByClassName("icnSearch")
    Dim isget As Object, targetlink As Object
    For Each isget In IE.Document.getElementsByClassName("icnSearch")
        '〇〇であればクリック
        If InStr(isget.outerHTML, "〇〇") > 0 Then
            Set targetlink = isget.parentElement.Children(1).Children(0)
            Exit For
            'targetlink.Click
            'Call secondClick
        ElseIf isget.innerText = "" And isget.parentElement.ClassName = "unread" Then
            Set targetlink = isget.parentElement.Children(1).Children(0)
            Exit For
            'targetlink.Click
            'Call secondClick
        End If
    Next
    
    If targetlink Is Nothing Then
        Call finish
    Else
        targetlink.Click
        Call secondClick
    End If

End Sub

Private Sub secondClick()

    Call lib.IE_Wait(IE)

    'エレメントを取得
    On Error GoTo Nolinks
    Dim pointlink As Object
    Set pointlink = IE.Document.getElementsByClassName("target_url")(0).getElementsByTagName("a")(0)

    'クリック
    pointlink.Click
    Call lib.IE_Wait(IE)

    '再帰してトップページへ
    Call Top
    Exit Sub

Nolinks:
    
    'Dim el As IHTMLElement
    Dim el As Object
        For Each el In IE.Document.links
            If InStr(el.href, "pmrd.rakuten.co.jp") > 0 Then
                el.Click
                Call lib.IE_Wait(IE)
                'Call IEStart(el.href)
                Exit For
            End If
        Next

    '再帰してトップページへ
    Call Top

End Sub

Private Sub finish()
        With CreateObject("WScript.Shell").popup("チェック完了", 2, "確認完了", vbInformation): End With
        IE.Quit
        Set IE = Nothing
        Sleep (10000): Set IE = Nothing: Set IE2 = Nothing: lib.AllcloseIEWindow
End Sub

