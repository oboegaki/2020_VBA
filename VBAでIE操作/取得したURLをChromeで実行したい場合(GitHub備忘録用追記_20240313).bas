Attribute VB_Name = "S004_PT_PointClick"
Option Explicit


Sub run()
    lib.IE_ImageOff
    
    '���O�C��
    Dim loginbtn As Object
    lib.IE_OPEN "https://�Z�Z", IE
    For Each loginbtn In IE.Document.getElementsByTagName("a")
        If InStr(loginbtn.outerHTML, "���O�C�����Ă�������") > 0 Then
            loginbtn.Click
        End If
    Next

    '�ړI�̏����Z���ɓ���
    With ThisWorkbook.Sheets("�Z�Z")
        If .Range("B3") = 0 Then
            Call pointnews("cat02")
            Call pointnews("cat03")
            .Range("B3") = 1
        End If
    End With

    'Chrome�ŊJ��
    Dim target_links As Object
    For Each target In IE.Document.getElementsByClassName("�Z�Z")
        For Each target_links In target.Children
            Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe " + target_links.Children(0).href
            Sleep 20000
        Next
    Next

    'Chrome�����
    Sleep 20000
    Shell "taskkill /F /IM chrome.exe"
    
    IE.Quit: Sleep (10000): Set IE = Nothing: lib.AllcloseIEWindow
    Sleep (10000): Set IE = Nothing: Set IE2 = Nothing: lib.AllcloseIEWindow
End Sub
