Attribute VB_Name = "S002_mail_de_point"
Option Explicit


Sub run()
    Application.StatusBar = "�N����..."
    lib.IE_ImageOff

    Call Top
End Sub

Private Sub Top()
        
    'visible��IE��S�ĕ���
    Call lib.AllcloseIEWindow
    lib.IE_OPEN "https://�Z�Z", IE
    
    'javascript���o��܂Œ��߂ɑҋ@
    Application.Wait Now + TimeValue("0:00:15")
    
    '�`�F�b�N
    'Dim targetlist: Set targetlist = IE.Document.getElementsByClassName("icnSearch")
    Dim isget As Object, targetlink As Object
    For Each isget In IE.Document.getElementsByClassName("icnSearch")
        '�Z�Z�ł���΃N���b�N
        If InStr(isget.outerHTML, "�Z�Z") > 0 Then
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

    '�G�������g���擾
    On Error GoTo Nolinks
    Dim pointlink As Object
    Set pointlink = IE.Document.getElementsByClassName("target_url")(0).getElementsByTagName("a")(0)

    '�N���b�N
    pointlink.Click
    Call lib.IE_Wait(IE)

    '�ċA���ăg�b�v�y�[�W��
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

    '�ċA���ăg�b�v�y�[�W��
    Call Top

End Sub

Private Sub finish()
        With CreateObject("WScript.Shell").popup("�`�F�b�N����", 2, "�m�F����", vbInformation): End With
        IE.Quit
        Set IE = Nothing
        Sleep (10000): Set IE = Nothing: Set IE2 = Nothing: lib.AllcloseIEWindow
End Sub

