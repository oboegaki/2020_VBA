Attribute VB_Name = "S004_PT_PointClick"
Option Explicit


Sub run()
    lib.IE_ImageOff
    
    'ログイン
    Dim loginbtn As Object
    lib.IE_OPEN "https://〇〇", IE
    For Each loginbtn In IE.Document.getElementsByTagName("a")
        If InStr(loginbtn.outerHTML, "ログインしてください") > 0 Then
            loginbtn.Click
        End If
    Next

    '目的の情報をセルに入力
    With ThisWorkbook.Sheets("〇〇")
        If .Range("B3") = 0 Then
            Call pointnews("cat02")
            Call pointnews("cat03")
            .Range("B3") = 1
        End If
    End With

    'Chromeで開く
    Dim target_links As Object
    For Each target In IE.Document.getElementsByClassName("〇〇")
        For Each target_links In target.Children
            Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe " + target_links.Children(0).href
            Sleep 20000
        Next
    Next

    'Chromeを閉じる
    Sleep 20000
    Shell "taskkill /F /IM chrome.exe"
    
    IE.Quit: Sleep (10000): Set IE = Nothing: lib.AllcloseIEWindow
    Sleep (10000): Set IE = Nothing: Set IE2 = Nothing: lib.AllcloseIEWindow
End Sub
