Attribute VB_Name = "Module7"
Option Explicit

'�����I�ɍőO�ʂɂ�����
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'���̃T�C�Y�ɖ߂�
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'public�p�錾
Dim IE As Object
Dim urls As Collection

Sub run()
    'initialize
    Set IE = CreateObject("InternetExplorer.Application")
    
    '�J��url���i�[
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
'STEP1 : IE�N��
'----------------------------------------------------------
Private Sub IEStart(ByVal url As String)
    IE.Visible = True
    IE.Navigate url

    '�ŏ�������Ă���ꍇ�͌��ɖ߂�(9=RESTORE:�ŏ����O�̏��)
    If IsIconic(IE.hWnd) Then
        ShowWindowAsync IE.hWnd, &H9
    End If
    SetForegroundWindow (IE.hWnd)     '�őO�ʂɕ\��
    
    Call loading
End Sub

Private Sub loading()
    '�T�C�g�̓ǂݍ��ݑ҂�
    Do While IE.Busy = True Or IE.ReadyState <> 4
        DoEvents
    Loop
    
    'JavaScript�ǂݍ��ݑ҂�
    Application.Wait Now + TimeValue("0:00:03")
End Sub

Private Sub IErefresh()
    IE.Refresh
    'JavaScript�ǂݍ��ݑ҂�
    Application.Wait Now + TimeValue("0:00:03")
End Sub

Private Sub entry()
        
    '�m�F�y�[�W��
    Call IEStart("https://�Z�Z")
    
    '�y�[�W�X�V
    Call IErefresh
    
    '���Ǎς݂łȂ���΃N���b�N
    Dim entryBtn As Object
    For Each entryBtn In IE.Document.getElementsByClassName("ui-btn-a")
        If "���Ǎς�" <> entryBtn.innerText Then
            entryBtn.Click
            Call loading
            Call answer
            Exit For
        End If
    Next
    
End Sub

Private Sub answer()
    
    '�Ăэ��ݑ҂� ��Call loading�͑ҋ@����������
    Application.Wait Now + TimeValue("0:00:10")

    '�I���y�[�W�ł���ΏI��
    If "https://�Z�Z" = IE.LocationURL Then: Call toContinue
            
    'a�^�O�΍�
    Dim linkend As Object
    If Not IsNull(IE.Document.getElementById("endlink")) Then
        IE.Document.getElementById("endlink").Click
        Call answer
    End If
    
    'form�擾
    Dim form As Object, forms As Object
    Set forms = IE.Document.getElementsByTagName("form")
    
    For Each form In forms
    
        'input�^�O�̃`�F�b�N
        Dim inputs As Object, targetinput As Object
        Set inputs = form.Document.getElementsByTagName("input")
            For Each targetinput In inputs
            
            'input�^�O�����W�I�ł����checked
             If targetinput.Type = "radio" Then
                 targetinput.Checked = True                              '�`�F�b�N
                 Exit For
             End If
        Next
        
        '�`�F�b�N�{�b�N�X��I��
        On Error Resume Next
        IE.Document.getElementById("que3").Checked = True
    
        'input�^�O�́ualt�v���u�i�ށv�̏ꍇ�͑J��
        If InStr(form.innerHTML, "�i��") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "�I��") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "�R������ǂ�") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "next") > 0 Then GoTo nextbtn
        If InStr(form.innerHTML, "�N���b�N���ă|�C���g�Q�b�g") > 0 Then GoTo nextbtn
    Next

    
    For Each form In forms
        Dim i As Long
        For i = 1 To urls.count
            If InStr(form.Action, urls.Item(i)) > 0 Then
                '�ړI��form����
                GoTo nextbtn
            End If
        Next i
    Next
    

    'form�o�Ȃ��ꍇ�A�I�������N��T��
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
    
    '���M���ċA
    Application.Wait Now + TimeValue("0:00:03")   '���b�҂�
    
    form.submit
    Set inputs = Nothing
    Set form = Nothing
    
    Set urls = Nothing '���Ƃŏ���
    
    Call answer
    
End Sub


Private Sub popup()
    Dim Obj As Object
    Dim re As Long
    Set Obj = CreateObject("WScript.Shell")
    re = Obj.popup("�����N���܂����f�o�b�O���܂�", 0, "�m�F", vbCritical)
    Stop
    Set Obj = Nothing
End Sub

Sub toContinue()
    Dim Obj As Object, re As Long
    Set Obj = CreateObject("WScript.Shell")
    re = Obj.popup("��Ƃ𑱂��܂���", 0, "����", vbYesNo + vbInformation)
    If re = vbYes Then
        IE.Quit
        Call run
    Else
        End
    End If
    Set Obj = Nothing
End Sub
